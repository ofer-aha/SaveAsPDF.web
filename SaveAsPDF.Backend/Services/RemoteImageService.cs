using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;

/// <summary>
/// Downloads the remote http(s) images an e-mail body points at and rewrites the
/// tags to carry the bytes inline as data: URIs, before the HTML ever reaches
/// Chromium.
///
/// Why this exists: modern signature blocks (WiseStamp, Exclaimer and friends)
/// do not embed their logos as cid: attachments - they link them from a CDN. The
/// renderer aborts every remote request when AllowRemoteContent is off, so those
/// signatures printed as bare alt text ("Photo", "Icon", "Linkedin Icon").
///
/// Why not simply let Chromium fetch them: an unreachable host leaves the request
/// pending, Networkidle0 never settles, and the render dies at the navigation
/// timeout and falls back to an HTML file - the exact failure the request blocking
/// was added to prevent. Fetching here instead keeps that guarantee: every fetch
/// has a hard per-image and total deadline, anything slow or unreachable is simply
/// dropped, and the page Chromium renders is still fully self-contained.
/// </summary>
public static class RemoteImageService
{
    // Ceilings, not expectations. A signature block is a handful of small logos;
    // anything past these numbers is a mail designed to make the server work.
    private const int  MaxImages        = 40;
    private const long MaxBytesPerImage = 2L * 1024 * 1024;   // 2 MB
    private const long MaxTotalBytes    = 8L * 1024 * 1024;   // 8 MB per render
    private const int  MaxRedirects     = 3;

    // src="http://..." / src='https://...' on an <img> tag. Deliberately narrow:
    // only <img> is inlined, so a mail cannot make the server pull stylesheets,
    // fonts, frames or scripts.
    private static readonly Regex ImgSrc = new(
        @"(<img\b[^>]*?\bsrc\s*=\s*)(""|')(https?://[^""'\s>]+)\2",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly HttpClient Http = CreateClient();

    private static HttpClient CreateClient()
    {
        // Redirects are followed by hand (see FetchAsync) so every hop can be
        // re-checked against the address filter - an automatic redirect is a free
        // ride from a public CDN to an internal host.
        var client = new HttpClient(new HttpClientHandler
        {
            AllowAutoRedirect      = false,
            AutomaticDecompression = DecompressionMethods.All
        });
        client.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0 (compatible; SaveAsPDF)");
        client.DefaultRequestHeaders.Accept.ParseAdd("image/*");
        return client;
    }

    /// <summary>
    /// Returns the HTML with every reachable remote image replaced by a data: URI.
    /// Never throws: an image that cannot be fetched keeps its original src, which
    /// the renderer then aborts exactly as before.
    /// </summary>
    public static async Task<string> InlineImagesAsync(
        string? html, int perImageTimeoutMs, int totalBudgetMs)
    {
        if (string.IsNullOrEmpty(html)) return html ?? "";

        var matches = ImgSrc.Matches(html);
        if (matches.Count == 0) return html;

        var urls = matches.Select(m => m.Groups[3].Value)
                          .Distinct(StringComparer.OrdinalIgnoreCase)
                          .Take(MaxImages)
                          .ToList();

        using var budget = new CancellationTokenSource(Math.Max(1000, totalBudgetMs));

        var fetched = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        long totalBytes = 0;
        var  ledger     = new object();

        var jobs = urls.Select(async url =>
        {
            var image = await FetchAsync(url, perImageTimeoutMs, budget.Token);
            if (image == null) return;

            lock (ledger)
            {
                if (totalBytes + image.Value.Bytes.LongLength > MaxTotalBytes) return;
                totalBytes += image.Value.Bytes.LongLength;
                fetched[url] =
                    $"data:{image.Value.ContentType};base64,{Convert.ToBase64String(image.Value.Bytes)}";
            }
        });

        try { await Task.WhenAll(jobs); } catch { /* every job swallows its own errors */ }

        if (fetched.Count == 0) return html;

        return ImgSrc.Replace(html, m =>
            fetched.TryGetValue(m.Groups[3].Value, out var dataUri)
                ? m.Groups[1].Value + m.Groups[2].Value + dataUri + m.Groups[2].Value
                : m.Value);
    }

    private static async Task<(byte[] Bytes, string ContentType)?> FetchAsync(
        string url, int perImageTimeoutMs, CancellationToken budget)
    {
        try
        {
            using var deadline = CancellationTokenSource.CreateLinkedTokenSource(budget);
            deadline.CancelAfter(Math.Max(500, perImageTimeoutMs));

            var current = new Uri(url);

            for (var hop = 0; hop <= MaxRedirects; hop++)
            {
                if (!await IsPubliclyRoutableAsync(current, deadline.Token)) return null;

                using var request  = new HttpRequestMessage(HttpMethod.Get, current);
                using var response = await Http.SendAsync(
                    request, HttpCompletionOption.ResponseHeadersRead, deadline.Token);

                if ((int)response.StatusCode is >= 300 and < 400 &&
                    response.Headers.Location != null)
                {
                    current = new Uri(current, response.Headers.Location);
                    if (current.Scheme != Uri.UriSchemeHttp && current.Scheme != Uri.UriSchemeHttps)
                        return null;
                    continue;
                }

                if (!response.IsSuccessStatusCode) return null;

                var contentType = response.Content.Headers.ContentType?.MediaType ?? "";
                if (!contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
                    return null;

                // Trust the declared length when there is one, but still cap the read:
                // Content-Length is the sender's claim, not a fact.
                if (response.Content.Headers.ContentLength > MaxBytesPerImage) return null;

                var bytes = await ReadCappedAsync(response, deadline.Token);
                return bytes is { LongLength: > 0 } ? (bytes, contentType) : null;
            }

            return null;   // redirect limit
        }
        catch
        {
            // Unreachable host, DNS failure, TLS rejection, timeout, oversized body.
            // A missing signature logo is never worth failing a save over.
            return null;
        }
    }

    private static async Task<byte[]?> ReadCappedAsync(
        HttpResponseMessage response, CancellationToken token)
    {
        await using var stream = await response.Content.ReadAsStreamAsync(token);
        using var buffer = new MemoryStream();

        var chunk = new byte[16 * 1024];
        int read;
        while ((read = await stream.ReadAsync(chunk, token)) > 0)
        {
            if (buffer.Length + read > MaxBytesPerImage) return null;
            buffer.Write(chunk, 0, read);
        }
        return buffer.ToArray();
    }

    // SSRF guard. The body of an e-mail is chosen by whoever sent it, so without
    // this an outsider can aim the file server at any host it can reach - an
    // intranet admin page, a cloud metadata endpoint, a printer. Only addresses
    // that are routable on the public internet are allowed.
    //
    // This resolves the name and then lets HttpClient resolve it again, so a
    // deliberately rebinding DNS record could in principle slip between the two
    // lookups. Pinning the connection to the verified address would close that
    // window; it is not done here because the payoff (fetching a decorative logo)
    // does not justify a custom SocketsHttpHandler connect callback.
    private static async Task<bool> IsPubliclyRoutableAsync(Uri uri, CancellationToken token)
    {
        if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) return false;

        IPAddress[] addresses;
        if (IPAddress.TryParse(uri.Host, out var literal)) addresses = [literal];
        else
        {
            try { addresses = await Dns.GetHostAddressesAsync(uri.Host, token); }
            catch { return false; }
        }

        return addresses.Length > 0 && addresses.All(a => !IsBlocked(a));
    }

    private static bool IsBlocked(IPAddress ip)
    {
        if (IPAddress.IsLoopback(ip)) return true;

        if (ip.AddressFamily == AddressFamily.InterNetworkV6)
        {
            if (ip.IsIPv4MappedToIPv6) return IsBlocked(ip.MapToIPv4());
            if (ip.IsIPv6LinkLocal || ip.IsIPv6SiteLocal || ip.IsIPv6Multicast) return true;
            if (IPAddress.IPv6Any.Equals(ip)) return true;
            var v6 = ip.GetAddressBytes();
            return v6[0] is 0xfc or 0xfd;                        // unique local fc00::/7
        }

        var b = ip.GetAddressBytes();
        return b[0] switch
        {
            0   => true,                                          // 0.0.0.0/8
            10  => true,                                          // private
            127 => true,                                          // loopback
            169 => b[1] == 254,                                    // link-local + metadata
            172 => b[1] >= 16 && b[1] <= 31,                       // private
            192 => b[1] == 168,                                    // private
            >= 224 => true,                                        // multicast + reserved
            _   => false
        };
    }
}
