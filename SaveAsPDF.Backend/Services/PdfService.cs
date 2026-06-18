using System.Reflection;
using System.Text;
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
                Args     = new[]
                {
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
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

        var safeSubject = Sanitize(email?.Subject ?? "(no subject)");
        var fileName    = $"{timestamp}_{safeSubject}.pdf";
        var pdfPath     = Path.Combine(projectPath, fileName);

        var html   = BuildHtml(email, request);
        var result = new PdfResult { FileName = fileName, FullPath = pdfPath, PdfCreated = false };

        // Generate to a local temp file first — avoids UNC path quirks.
        var localTemp = Path.Combine(Path.GetTempPath(), $"saveaspdf_{Guid.NewGuid():N}.pdf");
        try
        {
            var browser = await GetBrowserAsync();
            await using var page = await browser.NewPageAsync();

            await page.SetContentAsync(html, new NavigationOptions
            {
                WaitUntil = new[] { WaitUntilNavigation.Networkidle0 }
            });

            var ps = pdfSettings ?? new PdfSettings();
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

            File.Copy(localTemp, pdfPath, overwrite: true);
            result.PdfCreated = true;
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

            var htmlFallback = Path.ChangeExtension(pdfPath, ".html");
            try { File.WriteAllText(htmlFallback, html, Encoding.UTF8); } catch { }
            result.FileName       = Path.GetFileName(htmlFallback);
            result.FullPath       = htmlFallback;
            result.FallbackReason = reason;
        }
        finally
        {
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
        "<span style=\"direction:rtl;\">נוצר באמצעות SaveAsPDF גרסה " +
        Esc(version) + "</span>" +
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
        sb.Append(@"<!DOCTYPE html>
<html lang=""he""><head><meta charset=""UTF-8"" />
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
        sb.Append("<div class=\"body\">").Append(email.BodyHtml ?? "").Append("</div>");
        sb.Append("</div>");
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string Sanitize(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
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
