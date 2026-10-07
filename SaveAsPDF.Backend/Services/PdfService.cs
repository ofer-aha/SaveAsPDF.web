using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PuppeteerSharp;
using PuppeteerSharp.Media;
using static PuppeteerSharp.Media.PaperFormat;

/// <summary>
/// Generates PDFs using a headless Chromium browser (via PuppeteerSharp).
///
/// Why Chromium instead of iText7:
///   iText7's HTML renderer ignores CSS BiDi properties (direction, unicode-bidi)
///   and treats every glyph as LTR, producing broken/mirrored Hebrew text.
///   Chromium implements the Unicode Bidirectional Algorithm natively, so mixed
///   Hebrew/English text renders correctly with no pre-processing required.
///
/// Chromium download:
///   PuppeteerSharp downloads a pinned Chromium revision to
///   %USERPROFILE%\.cache\puppeteer  on first use (~170 MB, once per machine).
///   Subsequent starts are instant — the binary is reused.
/// </summary>
public static class PdfService
{
    // ── Singleton browser ────────────────────────────────────────────────────
    // One Chromium process shared across all requests; each request gets its
    // own Page, which is fully isolated and safe for concurrent use.

    private static IBrowser?          _browser;
    private static readonly SemaphoreSlim _initLock = new(1, 1);

    // One render can hold a Chromium tab for the full render timeout (up to 120 s)
    // while the calling thread blocks on it. Without a cap, a burst of requests
    // opens a tab per request and pins a thread per request, taking the whole
    // service down. Four concurrent renders is comfortably more than the real
    // workload and bounds the damage.
    private const int MaxConcurrentRenders = 4;
    private static readonly SemaphoreSlim _renderGate = new(MaxConcurrentRenders, MaxConcurrentRenders);

    // Selected Chromium build (BrowserFetcher buildId). null = PuppeteerSharp's
    // default pinned build. Set from settings on startup and after an engine update.
    private static string? _preferredBuild;

    private static async Task<IBrowser> GetBrowserAsync()
    {
        if (_browser is { IsConnected: true }) return _browser;

        await _initLock.WaitAsync();
        try
        {
            if (_browser is { IsConnected: true }) return _browser;

            // Resolve the Chromium executable. With a preferred build, use (and if
            // needed download) exactly that build; otherwise fall back to the
            // default pinned build (~170 MB, downloaded once per machine).
            var fetcher = new BrowserFetcher();
            string? exePath = null;
            if (!string.IsNullOrWhiteSpace(_preferredBuild))
            {
                var inst = fetcher.GetInstalledBrowsers()
                                  .FirstOrDefault(b => b.BuildId == _preferredBuild)
                           ?? await fetcher.DownloadAsync(_preferredBuild);
                exePath = inst.GetExecutablePath();
            }
            else
            {
                await fetcher.DownloadAsync();
            }

            _browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless       = true,
                ExecutablePath = exePath,   // null → PuppeteerSharp default resolution
                // NOTE: --no-sandbox / --disable-setuid-sandbox were here. They are a
                // Linux-container workaround and are NOT needed for a Windows
                // service - and this renderer parses fully attacker-chosen HTML
                // (any e-mail body), so turning the OS sandbox off removed the one
                // thing containing a renderer bug.
                Args     = new[]
                {
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--font-render-hinting=none"   // crisper text at PDF resolution
                }
            });
            return _browser;
        }
        finally { _initLock.Release(); }
    }

    // ── Engine management (admin PDF tab) ─────────────────────────────────────

    /// <summary>PuppeteerSharp assembly version driving the PDF engine.</summary>
    public static string PuppeteerVersion =>
        typeof(Puppeteer).Assembly.GetName().Version?.ToString() ?? "unknown";

    /// <summary>Apply the configured Chromium build (called on startup).</summary>
    public static void SetPreferredBuild(string? buildId) =>
        _preferredBuild = string.IsNullOrWhiteSpace(buildId) ? null : buildId.Trim();

    /// <summary>Close the shared browser so the next request relaunches it
    /// (e.g. after switching to a freshly downloaded Chromium build).</summary>
    public static async Task DisposeBrowserAsync()
    {
        await _initLock.WaitAsync();
        try
        {
            if (_browser != null)
            {
                try { await _browser.CloseAsync(); } catch { }
                _browser = null;
            }
        }
        finally { _initLock.Release(); }
    }

    /// <summary>Current engine state for the admin UI.</summary>
    public static (string current, string[] installed, string cacheDir) GetEngineState()
    {
        var f = new BrowserFetcher();
        string[] installed;
        try
        {
            installed = f.GetInstalledBrowsers()
                         .Select(b => b.BuildId)
                         .Distinct()
                         .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                         .ToArray();
        }
        catch { installed = Array.Empty<string>(); }

        var current = _preferredBuild ?? installed.LastOrDefault() ?? "";
        return (current, installed, f.CacheDir);
    }

    /// <summary>Download the latest Chromium build and switch the engine to it.
    /// Returns the new buildId.</summary>
    public static async Task<string> UpdateToLatestAsync()
    {
        var f        = new BrowserFetcher();
        var installed = await f.DownloadAsync(BrowserTag.Latest);
        SetPreferredBuild(installed.BuildId);
        await DisposeBrowserAsync();   // next render uses the new build
        return installed.BuildId;
    }

    /// <summary>Switch the engine to an already-installed Chromium build, or to
    /// PuppeteerSharp's default pinned build when <paramref name="buildId"/> is
    /// null/empty. Downloads the build if it is not present locally.
    /// Returns the build now in use ("" = default pinned build).</summary>
    public static async Task<string> SelectBuildAsync(string? buildId)
    {
        var wanted = string.IsNullOrWhiteSpace(buildId) ? null : buildId.Trim();

        if (wanted != null)
        {
            var f = new BrowserFetcher();
            var have = f.GetInstalledBrowsers().Any(b => b.BuildId == wanted);
            if (!have) await f.DownloadAsync(wanted);
        }

        SetPreferredBuild(wanted);
        await DisposeBrowserAsync();   // next render uses the selected build
        return wanted ?? "";
    }

    // ── Public entry point ────────────────────────────────────────────────────

    public static PdfResult GeneratePdf(
        EmailDto?         email,
        string            projectPath,
        SaveAsPdfRequest? request     = null,
        PdfSettings?      pdfSettings = null)
    {
        // Controllers are synchronous; run the async work on the thread pool.
        return Task.Run(() => GeneratePdfAsync(email, projectPath, request, pdfSettings))
                   .GetAwaiter().GetResult();
    }

    private static async Task<PdfResult> GeneratePdfAsync(
        EmailDto?         email,
        string            projectPath,
        SaveAsPdfRequest? request     = null,
        PdfSettings?      pdfSettings = null)
    {
        var receivedDate = email?.ReceivedDate;
        var timestamp    = receivedDate?.ToString("yyyyMMdd_HHmm")
                           ?? DateTime.Now.ToString("yyyyMMdd_HHmm");

        // Outlook subjects run to 255 characters and folder names to 200, so an
        // unclamped name blows past MAX_PATH (260) and every write throws - which
        // used to surface as a save that silently produced nothing.
        var safeSubject = Sanitize(email?.Subject ?? "(no subject)");
        var pdfPath     = BuildUniquePath(projectPath, timestamp, safeSubject, ".pdf");
        var fileName    = Path.GetFileName(pdfPath);

        var html   = BuildHtml(email, request);
        var result = new PdfResult { FileName = fileName, FullPath = pdfPath, PdfCreated = false };

        // Generate to a local temp file first — avoids UNC path quirks.
        var localTemp = Path.Combine(Path.GetTempPath(), $"saveaspdf_{Guid.NewGuid():N}.pdf");
        var gateTaken = false;
        try
        {
            // Bound concurrent renders. Waiting is fine - the client is already
            // waiting on this request - but a caller that cannot get a slot within
            // two minutes falls through to the catch and gets the HTML fallback
            // rather than piling up behind the others.
            gateTaken = await _renderGate.WaitAsync(TimeSpan.FromSeconds(120));
            if (!gateTaken)
                throw new TimeoutException(
                    "PDF renderer is busy - too many concurrent conversions.");

            var browser = await GetBrowserAsync();
            await using var page = await browser.NewPageAsync();

            var ps = pdfSettings ?? new PdfSettings();
            var renderTimeout = Math.Clamp(ps.RenderTimeoutMs <= 0 ? 15000 : ps.RenderTimeoutMs,
                                           2000, 120000);

            // Signature logos are usually linked from a CDN rather than embedded as
            // cid: attachments, so with remote requests blocked they printed as bare
            // alt text. Fetch them here - on a leash - and hand Chromium a fully
            // self-contained document, keeping the "renderer never waits on the
            // network" guarantee that the request blocking below exists to provide.
            // Pointless when AllowRemoteContent already lets the page fetch its own.
            if (ps.AllowRemoteImages && !ps.AllowRemoteContent)
                html = await RemoteImageService.InlineImagesAsync(
                    html, ps.RemoteImageTimeoutMs, Math.Min(renderTimeout, 20000));

            // SECURITY: the email body is entirely attacker-controlled - anyone who can
            // send mail to a user can choose what this server renders. Scripts are never
            // needed to lay out an email, so turn them off. Without this, a crafted
            // message runs arbitrary JavaScript inside a Chromium started with
            // --no-sandbox, which is a far larger blast radius than a malformed PDF.
            await page.SetJavaScriptEnabledAsync(false);

            // Block remote content unless explicitly allowed.
            //
            // A remote image the server cannot reach does not fail fast - the request
            // sits pending, so Networkidle0 never settles and the render stalls until
            // the navigation timeout fires. Users saw that as
            // "[TimeoutException] Timeout of 30000 ms exceeded" plus an HTML fallback
            // instead of a PDF. Aborting external requests up front removes the entire
            // class of hang. Inline content (data:, cid:) is untouched, so embedded
            // images still render.
            if (!ps.AllowRemoteContent)
            {
                await page.SetRequestInterceptionAsync(true);
                page.Request += async (_, e) =>
                {
                    try
                    {
                        // Allow-list, not block-list. Naming http/https as "remote" let
                        // every other scheme through - file:, ftp:, chrome: included.
                        // Only schemes that carry content already inside this document
                        // are permitted; everything else is aborted.
                        var url = e.Request.Url ?? "";
                        var isInline =
                            url.StartsWith("data:",  StringComparison.OrdinalIgnoreCase) ||
                            url.StartsWith("cid:",   StringComparison.OrdinalIgnoreCase) ||
                            url.StartsWith("blob:",  StringComparison.OrdinalIgnoreCase) ||
                            url.StartsWith("about:", StringComparison.OrdinalIgnoreCase);
                        if (isInline) await e.Request.ContinueAsync();
                        else          await e.Request.AbortAsync();
                    }
                    catch { /* request already handled or page gone */ }
                };
            }

            // Fallback ladder. Each step is strictly weaker than the last, so a slow or
            // awkward email degrades to a slightly less settled PDF rather than to no
            // PDF at all. Dropping to HTML is now reserved for real failures.
            var waitStages = new[]
            {
                new[] { WaitUntilNavigation.Networkidle0 },
                new[] { WaitUntilNavigation.Load },
                new[] { WaitUntilNavigation.DOMContentLoaded }
            };

            Exception? lastNavError = null;
            var settled = false;
            for (var stage = 0; stage < waitStages.Length && !settled; stage++)
            {
                // Later stages get a shorter budget - by then we mainly want the DOM.
                var budget = stage == 0 ? renderTimeout : Math.Min(5000, renderTimeout);
                try
                {
                    await page.SetContentAsync(html, new NavigationOptions
                    {
                        WaitUntil = waitStages[stage],
                        Timeout   = budget
                    });
                    settled = true;
                }
                catch (Exception navEx)
                {
                    lastNavError = navEx;
                }
            }

            if (!settled && lastNavError != null)
            {
                // Content is in the DOM even when the wait condition never fired, so
                // printing anyway generally yields a usable PDF.
                Console.Error.WriteLine(
                    $"[PdfService] Page never settled ({lastNavError.GetType().Name}: " +
                    $"{lastNavError.Message}). Printing current state anyway.");
            }

            await page.PdfAsync(localTemp, new PdfOptions
            {
                Format              = ResolveFormat(ps.PageSize),
                Landscape           = ps.Landscape,
                PrintBackground     = ps.PrintBackground,
                DisplayHeaderFooter = true,
                HeaderTemplate      = BuildHeaderTemplate(email?.Subject),
                FooterTemplate      = BuildFooterTemplate(ReadAppVersion()),
                MarginOptions   = new MarginOptions
                {
                    Top    = $"{ps.MarginTopCm:F2}cm",
                    Bottom = $"{ps.MarginBottomCm:F2}cm",
                    Left   = $"{ps.MarginLeftCm:F2}cm",
                    Right  = $"{ps.MarginRightCm:F2}cm"
                }
            });

            // overwrite:false - BuildUniquePath already picked a free name, and a
            // collision appearing between then and now means another save won the
            // race; silently flattening it would destroy that e-mail.
            File.Copy(localTemp, pdfPath, overwrite: false);
            result.PdfCreated = true;
            result.FileWritten = true;
        }
        catch (Exception ex)
        {
            var msg = new StringBuilder();
            for (var e = ex; e != null; e = e.InnerException)
                msg.Append(e == ex ? "" : " → ")
                   .Append('[').Append(e.GetType().Name).Append("] ")
                   .Append(e.Message);
            var reason = msg.ToString();
            if (reason.Length > 800) reason = reason[..800] + "…";

            // The fallback can fail for exactly the same reasons the PDF did
            // (share offline, read-only, disk full, path too long). Reporting a
            // file name regardless is how a save that wrote nothing still showed
            // up as "saved" and got the e-mail marked as archived.
            var htmlFallback = Path.ChangeExtension(pdfPath, ".html");
            try
            {
                File.WriteAllText(htmlFallback, html, Encoding.UTF8);
                result.FileName    = Path.GetFileName(htmlFallback);
                result.FullPath    = htmlFallback;
                result.FileWritten = true;
            }
            catch (Exception fallbackEx)
            {
                result.FileWritten = false;
                reason += $" | HTML fallback also failed: [{fallbackEx.GetType().Name}] {fallbackEx.Message}";
            }
            result.FallbackReason = reason;
        }
        finally
        {
            if (gateTaken) _renderGate.Release();
            try { File.Delete(localTemp); } catch { }
        }

        return result;
    }

    // ── Helpers ── (PDF options) ──────────────────────────────────────────────

    private static PaperFormat ResolveFormat(string? size) => size switch
    {
        "Letter"  => Letter,
        "Legal"   => Legal,
        "A3"      => A3,
        _         => A4     // default
    };

    // ── Header / footer templates ─────────────────────────────────────────────
    // Chromium renders these as standalone HTML fragments. Note: the default
    // font-size in header/footer templates is 0, so an explicit font-size is
    // required or nothing shows. Special spans (pageNumber/totalPages) are
    // substituted by Chromium at render time.

    private static string BuildHeaderTemplate(string? subject) =>
        "<div style=\"font-size:9px; width:100%; padding:0 1.2cm; " +
        "color:#666; text-align:center; direction:rtl;\" dir=\"auto\">" +
        "<span style=\"font-weight:600;\">Subject:</span> " +
        Esc(subject ?? "(no subject)") +
        "</div>";

    private static string BuildFooterTemplate(string version) =>
        "<div style=\"font-size:9px; width:100%; padding:0 1.2cm; color:#666; " +
        "display:flex; justify-content:space-between; align-items:center;\">" +
        "<span style=\"direction:ltr;\">Page <span class=\"pageNumber\"></span> " +
        "of <span class=\"totalPages\"></span></span>" +
        // Credit line. The two addresses sit in their own dir=ltr span: dropped
        // straight into the RTL run they are LTR text separated by a neutral
        // space, which the bidi algorithm is free to reorder. nowrap keeps the
        // line on one row - Chromium clips a footer that grows past the bottom
        // margin rather than expanding it.
        "<span style=\"direction:rtl; white-space:nowrap;\">" +
        "נוצר באמצעות SaveAsPDF " + Esc(version) + " תכנון וביצוע: עופר אהרון " +
        "<span style=\"direction:ltr; unicode-bidi:embed;\">" +
        "ofer@sw-eng.co.il ofer.aha@gmail.com</span>" +
        "</span>" +
        "</div>";

    // Reads the app version from package.json deployed alongside the binary —
    // the same single source of truth used by InfoController (/api/info).
    private static string ReadAppVersion()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, "package.json");
            if (File.Exists(path))
            {
                using var doc = System.Text.Json.JsonDocument.Parse(File.ReadAllText(path));
                var v = doc.RootElement.GetProperty("version").GetString();
                if (!string.IsNullOrWhiteSpace(v)) return v!;
            }
        }
        catch { /* fall through to assembly version */ }

        return typeof(PdfService).Assembly
                   .GetCustomAttribute<System.Reflection.AssemblyInformationalVersionAttribute>()
                   ?.InformationalVersion?.Split('+')[0]
               ?? "0.0.0";
    }

    // ── HTML composition ──────────────────────────────────────────────────────
    // Chromium implements the Unicode BiDi Algorithm and respects CSS
    // direction/unicode-bidi, so no character-level RTL pre-processing is needed.

    private static string BuildHtml(EmailDto? email, SaveAsPdfRequest? req)
    {
        var sb = new StringBuilder();
        // The CSP matters most OUTSIDE the renderer: when a render fails this same
        // HTML is written next to the PDFs as a .html file on the share, where no
        // SetJavaScriptEnabledAsync(false) applies. A user double-clicking it would
        // otherwise run whatever script the e-mail (or a stamp template) contained.
        sb.Append(@"<!DOCTYPE html>
<html lang=""he""><head><meta charset=""UTF-8"" />
<meta http-equiv=""Content-Security-Policy"" content=""default-src 'none'; img-src data: cid: blob:; style-src 'unsafe-inline'; font-src data:; base-uri 'none'; form-action 'none'"" />
<style>
  body {
    font-family: Arial, 'Segoe UI', Tahoma, Verdana, sans-serif;
    color: #222; line-height: 1.5; font-size: 12px; margin: 0; padding: 0;
  }
  .saveaspdf-stamp {
    direction: rtl; text-align: right;
    border: 2px solid #0078d4; background: #f0f6fc; border-radius: 6px;
    padding: 12px 16px; margin-bottom: 18px;
  }
  .saveaspdf-stamp table { border-collapse: collapse; width: 100%; }
  .saveaspdf-stamp td    { padding: 2px 0 2px 8px; vertical-align: top; font-size: 12px; }
  .saveaspdf-stamp .label { color: #666; white-space: nowrap; width: 1%; }
  .saveaspdf-stamp .fwd   { color: #107c10; font-weight: 600; }
  .saveaspdf-stamp .msg-header { border-top: 1px solid #cfe0f0; margin-top: 12px; padding-top: 10px; }
  .saveaspdf-stamp .msg-header-title {
    color: #0078d4; font-weight: 700; font-size: 12px;
    margin-bottom: 6px; letter-spacing: .2px;
  }
  .saveaspdf-stamp .msg-header table { border-collapse: collapse; width: 100%; }
  .saveaspdf-stamp .msg-header td { padding: 3px 0 3px 10px; vertical-align: top; font-size: 12px; }
  .saveaspdf-stamp .msg-header td.label {
    color: #555; font-weight: 600; white-space: nowrap;
    width: 72px; text-align: right;
  }
  .saveaspdf-stamp .msg-header td.value { color: #222; word-break: break-word; }
  .original { border-top: 1px solid #ddd; padding-top: 14px; }
  .original .meta { font-size: 11px; color: #666; margin-bottom: 6px;
                    direction: rtl; text-align: right; }
  .original h3    { margin: 0 0 8px 0; font-size: 14px;
                    direction: rtl; text-align: right; }
  .original .body { direction: rtl; unicode-bidi: embed; margin-top: 10px; }
  /* Outlook/Gmail sometimes emit ol/ul with an inline display:flex, which
     Chromium honors by laying list items out side-by-side as columns
     (garbled in the PDF, though Outlook shows them stacked). Force lists
     back to normal vertical flow; !important beats the inline style. */
  .original .body ol,
  .original .body ul   { display: block !important; }
  .original .body li   { display: list-item !important; }
</style>
</head><body>");

        AppendStamp(sb, req, email);
        AppendOriginalEmail(sb, email);

        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static void AppendStamp(StringBuilder sb, SaveAsPdfRequest? req, EmailDto? email = null)
    {
        var stamp = req?.Stamp;

        // No stamp at all (user opted out, or admin forced stamping off) → no frame.
        if (stamp == null) return;

        sb.Append("<div class=\"saveaspdf-stamp\">");

        if (!string.IsNullOrWhiteSpace(stamp.Template))
        {
            var rendered = stamp.Template
                .Replace("{{projectId}}",   Esc(req?.ProjectId))
                .Replace("{{projectName}}", Esc(req?.ProjectName))
                .Replace("{{leader}}",      Esc(req?.ProjectLeader))
                .Replace("{{date}}",        Esc(DateTime.Now.ToString("dd/MM/yyyy HH:mm")))
                .Replace("{{user}}",        Esc(stamp.UserName))
                .Replace("{{employees}}",   FormatEmployees(req?.Employees))
                .Replace("{{attachments}}", FormatAttachmentNames(stamp.AttachmentNames))
                .Replace("{{notes}}",       Esc(stamp.Notes));
            sb.Append(rendered);
            AppendMessageHeader(sb, email, stamp);
            sb.Append("</div>");
            return;
        }

        sb.Append("<table>");

        void Row(string label, string value) =>
            sb.Append("<tr><td class=\"label\">").Append(label)
              .Append("</td><td>").Append(value).Append("</td></tr>");

        var empty = "<i style=\"color:#999\">(אין)</i>";

        if (stamp.IncludeProjectId)
            Row("מספר פרויקט:",
                !string.IsNullOrWhiteSpace(req?.ProjectId)
                    ? $"<b>{Esc(req.ProjectId)}</b>"
                    : empty);

        if (stamp.IncludeProjectName)
            Row("שם פרויקט:",
                !string.IsNullOrWhiteSpace(req?.ProjectName)
                    ? Esc(req.ProjectName)
                    : empty);

        if (stamp.IncludeLeader)
        {
            var leaderEmployee = req?.Employees?.FirstOrDefault(e => e.IsLeader);
            var leaderDisplay  = leaderEmployee?.DisplayName?.Trim();
            var leaderEmail    = leaderEmployee?.Email?.Trim();
            if (string.IsNullOrWhiteSpace(leaderDisplay))
                leaderDisplay = req?.ProjectLeader?.Trim();
            string leaderCell;
            if (string.IsNullOrWhiteSpace(leaderDisplay))
            {
                leaderCell = empty;
            }
            else
            {
                leaderCell = !string.IsNullOrWhiteSpace(leaderEmail)
                    ? $"<a href=\"mailto:{Esc(leaderEmail)}\">{Esc(leaderDisplay)}</a>"
                    : Esc(leaderDisplay);
            }
            Row("מנהל פרויקט:", leaderCell);
        }

        if (stamp.IncludeDate)
            Row("תאריך שמירה:", Esc(DateTime.Now.ToString("dd/MM/yyyy HH:mm")));

        if (stamp.IncludeUser)
            Row("נשמר על-ידי:",
                string.IsNullOrWhiteSpace(stamp.UserName) ? empty : Esc(stamp.UserName));

        if (stamp.IncludeEmployees)
            Row("עובדי פרויקט:",
                req?.Employees?.Count > 0 ? FormatEmployees(req.Employees) : empty);

        if (stamp.IncludeAttachments)
            Row("קבצים מצורפים:",
                stamp.AttachmentNames?.Count > 0 ? FormatAttachmentNames(stamp.AttachmentNames) : empty);

        if (stamp.Forwarded)
        {
            var fwdVal = "כן";
            if (!string.IsNullOrWhiteSpace(stamp.ForwardedTo))
                fwdVal += $" ({Esc(stamp.ForwardedTo)})";
            sb.Append("<tr><td class=\"label\">הועבר למנהל:</td>")
              .Append("<td class=\"fwd\">").Append(fwdVal).Append("</td></tr>");
        }

        if (!string.IsNullOrWhiteSpace(stamp.Notes))
            Row("הערות:", Esc(stamp.Notes));

        sb.Append("</table>");

        if (stamp.PolicyApplied)
            sb.Append("<div style=\"margin-top:8px;font-size:10px;color:#888;font-style:italic\">")
              .Append("[נעול] חלק מהשדות נקבעו על-ידי מנהל המערכת</div>");

        // Message header (from/to/sent/received…) rendered inside the same frame
        AppendMessageHeader(sb, email, stamp);

        sb.Append("</div>");
    }

    // Renders the original message's header fields (subject, from, to, cc, sent,
    // received) as a sub-block inside the SaveAsPDF frame, so the reader sees the
    // message's provenance and where it sits alongside the save details.
    private static void AppendMessageHeader(StringBuilder sb, EmailDto? email, StampInfo? stamp)
    {
        if (email == null) return;

        // Each header field is shown when its stamp toggle is on. When there is no
        // stamp config at all (stamp == null), every available field is shown.
        bool On(bool flag) => stamp == null || flag;

        var rows = new StringBuilder();
        void HRow(string label, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            rows.Append("<tr><td class=\"label\">").Append(label)
                .Append("</td><td class=\"value\">").Append(value).Append("</td></tr>");
        }

        // Clean, fixed field order. Labels align in their own column (see CSS).
        if (On(stamp?.IncludeFrom ?? true))
            HRow("מאת:",   Esc(email.From));                                     // From
        if (On(stamp?.IncludeTo ?? true) && email.To?.Count > 0)
            HRow("אל:",    Esc(string.Join(", ", email.To)));                    // To
        if (On(stamp?.IncludeCc ?? true) && email.Cc?.Count > 0)
            HRow("עותק:",  Esc(string.Join(", ", email.Cc)));                    // CC
        if (On(stamp?.IncludeSent ?? true) && email.SentDate.HasValue)
            HRow("נשלח:",  Esc(email.SentDate.Value.ToString("dd/MM/yyyy HH:mm")));  // Sent
        if (On(stamp?.IncludeReceived ?? true) && email.ReceivedDate.HasValue)
            HRow("התקבל:", Esc(email.ReceivedDate.Value.ToString("dd/MM/yyyy HH:mm"))); // Received
        if (On(stamp?.IncludeSubject ?? true))
            HRow("נושא:",  Esc(email.Subject));                                  // Subject

        if (rows.Length == 0) return;

        sb.Append("<div class=\"msg-header\">")
          .Append("<div class=\"msg-header-title\">פרטי ההודעה</div>")
          .Append("<table>").Append(rows).Append("</table></div>");
    }

    private static void AppendOriginalEmail(StringBuilder sb, EmailDto? email)
    {
        if (email == null) return;
        sb.Append("<div class=\"original\">");
        sb.Append("<h3>").Append(Esc(email.Subject ?? "(ללא נושא)")).Append("</h3>");

        // From/To/Cc/dates now live in the SaveAsPDF frame above (AppendMessageHeader);
        // this section holds only the subject heading and the verbatim body.

        // Emit the email body verbatim — Chromium handles RTL/bidi natively via CSS
        sb.Append("<div class=\"body\">").Append(SanitizeBodyHtml(email.BodyHtml)).Append("</div>");
        sb.Append("</div>");
    }

    // Matches a src="" / src='' attribute (optionally whitespace-only) on any
    // tag. Outlook leaves these behind for images it didn't download — e.g.
    // "Pictures from this sender weren't automatically downloaded" placeholders
    // for blocked external/remote signature logos and banners — while keeping
    // the original width/height.
    private static readonly Regex EmptySrcAttr =
        new(@"\ssrc\s*=\s*(""\s*""|'\s*')", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // An empty src is not a no-op: per the HTML spec, browsers (Chromium
    // included) resolve src="" to the *current document's own URL* and issue a
    // request for it. Inside page.SetContentAsync(..., WaitUntilNavigation.Networkidle0)
    // that bogus request can keep the connection count above zero indefinitely,
    // so navigation never settles — silently hanging the render until
    // PuppeteerSharp's 30s default navigation timeout fires. That is the exact
    // "[TimeoutException] Timeout of 30000 ms exceeded" users see as an
    // HTML-fallback. Stripping the empty attribute lets a broken image collapse
    // to nothing instead of stalling the whole PDF.
    private static string SanitizeBodyHtml(string? bodyHtml) =>
        string.IsNullOrEmpty(bodyHtml) ? "" : EmptySrcAttr.Replace(bodyHtml, "");

    // ── Helpers ───────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds "&lt;dir&gt;\&lt;timestamp&gt;_&lt;subject&gt;&lt;ext&gt;", shortening the subject so the
    /// whole path stays inside MAX_PATH, and appending " (2)", " (3)"... rather
    /// than overwriting an existing file.
    /// </summary>
    /// <remarks>
    /// Both halves fix real data loss. Unclamped, a long Hebrew subject inside a
    /// deep project folder pushed the path past 260 characters and every write
    /// threw. Un-uniqued, two e-mails with the same subject saved in the same
    /// minute - normal when working through a backlog of "דוח יומי" - silently
    /// overwrote each other, while their attachments were correctly kept apart.
    /// </remarks>
    private const int MaxFullPathLength = 240;   // MAX_PATH 260, minus headroom

    private static string BuildUniquePath(string dir, string timestamp, string subject, string ext)
    {
        // Longest suffix we may need to append for de-duplication.
        const string widestSuffix = " (99)";

        var prefix    = timestamp + "_";
        var overhead  = dir.Length + 1 + prefix.Length + ext.Length + widestSuffix.Length;
        var available = MaxFullPathLength - overhead;

        if (available < 8) available = 8;          // deep folder: keep a stub name
        if (subject.Length > available)
            subject = subject[..available].TrimEnd('.', ' ');
        if (subject.Length == 0) subject = "email";

        var candidate = Path.Combine(dir, prefix + subject + ext);
        if (!File.Exists(candidate)) return candidate;

        for (var i = 2; i <= 99; i++)
        {
            candidate = Path.Combine(dir, $"{prefix}{subject} ({i}){ext}");
            if (!File.Exists(candidate)) return candidate;
        }

        // 99 collisions in one minute is not a real workflow, but never overwrite.
        var unique = Guid.NewGuid().ToString("N")[..8];
        return Path.Combine(dir, $"{prefix}{subject} ({unique}){ext}");
    }

    private static string Sanitize(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        // Windows silently drops trailing dots and spaces, so a name ending in
        // one resolves to a different file than the one reported to the user.
        return name.Trim().TrimEnd('.', ' ');
    }

    private static string Esc(string? s) =>
        string.IsNullOrEmpty(s) ? "" :
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static string FormatEmployees(List<EmployeeDto>? employees)
    {
        if (employees == null || employees.Count == 0) return "";
        return string.Join(" ,", employees.Select(e =>
        {
            var label = Esc(e.DisplayName ?? e.Email ?? "");
            return e.IsLeader
                ? $"<b>{label}</b> <span style=\"color:#107c10;font-size:11px\">(מנהל)</span>"
                : label;
        }));
    }

    private static string FormatAttachmentNames(List<string>? names)
    {
        if (names == null || names.Count == 0) return "";
        return string.Join(", ", names
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(Esc));
    }
}
