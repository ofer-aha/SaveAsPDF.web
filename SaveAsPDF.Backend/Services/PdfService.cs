using System.Diagnostics;
using System.Text;

public static class PdfService
{
    // Probe order for headless-capable browsers
    private static readonly string[] BrowserCandidates =
    [
        @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        @"C:\Program Files\Google\Chrome\Application\chrome.exe",
        @"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    ];

    public static PdfResult GeneratePdf(
        EmailDto? email,
        string projectPath,
        SaveAsPdfRequest? request = null)
    {
        var receivedDate = email?.ReceivedDate;
        var timestamp    = receivedDate?.ToString("yyyyMMdd_HHmm")
                           ?? DateTime.Now.ToString("yyyyMMdd_HHmm");

        var safeSubject = Sanitize(email?.Subject ?? "(no subject)");
        var fileName    = $"{timestamp}_{safeSubject}.pdf";
        var pdfPath     = Path.Combine(projectPath, fileName);

        // Build the full HTML (stamp header + email metadata + original body)
        var html = BuildHtml(email, request);

        // Write to a temp file (Edge headless reads file:// URLs)
        var tempHtml = Path.Combine(Path.GetTempPath(), $"saveaspdf_{Guid.NewGuid():N}.html");
        File.WriteAllText(tempHtml, html, Encoding.UTF8);

        var result = new PdfResult { FileName = fileName, FullPath = pdfPath, PdfCreated = false };

        try
        {
            var (ok, reason) = HtmlToPdf(tempHtml, pdfPath);
            if (ok)
            {
                result.PdfCreated = true;
            }
            else
            {
                // Fallback: keep the HTML next to the would-be PDF so nothing is lost
                var htmlFallback = Path.ChangeExtension(pdfPath, ".html");
                File.Copy(tempHtml, htmlFallback, overwrite: true);
                result.FileName       = Path.GetFileName(htmlFallback);
                result.FullPath       = htmlFallback;
                result.FallbackReason = reason;
            }
        }
        finally
        {
            try { File.Delete(tempHtml); } catch { }
        }

        return result;
    }

    // ---------------------------------------------------------------
    // HTML composition
    // ---------------------------------------------------------------
    private static string BuildHtml(EmailDto? email, SaveAsPdfRequest? req)
    {
        var sb = new StringBuilder();
        sb.Append(@"<!DOCTYPE html>
<html dir=""rtl"" lang=""he""><head><meta charset=""UTF-8"" />
<style>
  @page { margin: 1.5cm; }
  body { font-family: Arial, 'Segoe UI', Tahoma, sans-serif; color: #222; line-height: 1.5; font-size: 12px; margin:0; }
  .saveaspdf-stamp {
    border: 2px solid #0078d4; background: #f0f6fc; border-radius: 6px;
    padding: 12px 16px; margin-bottom: 18px;
  }
  .saveaspdf-stamp h2 { margin: 0 0 8px 0; font-size: 14px; color: #0078d4; }
  .saveaspdf-stamp table { border-collapse: collapse; width: 100%; }
  .saveaspdf-stamp td { padding: 2px 8px 2px 0; vertical-align: top; font-size: 12px; }
  .saveaspdf-stamp .label { color: #666; white-space: nowrap; width: 1%; }
  .saveaspdf-stamp .fwd { color: #107c10; font-weight: 600; }
  .original { border-top: 1px solid #ddd; padding-top: 14px; }
  .original .meta { font-size: 11px; color: #666; margin-bottom: 6px; }
  .original h3 { margin: 0 0 8px 0; font-size: 14px; }
</style>
</head><body>");

        AppendStamp(sb, req);
        AppendOriginalEmail(sb, email);

        sb.Append("</body></html>");
        return sb.ToString();
    }

    private static void AppendStamp(StringBuilder sb, SaveAsPdfRequest? req)
    {
        var stamp = req?.Stamp;
        if (stamp == null) return;

        // Custom template overrides everything if provided.
        // All user-supplied placeholder VALUES are HTML-escaped before being
        // inserted, so a project name/leader/notes containing markup or
        // <script> can't break out of the placeholder slot.
        // FormatEmployees / FormatAttachmentNames produce already-escaped HTML
        // (formatted with badges) so they're inserted as-is.
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
            sb.Append("<div class=\"saveaspdf-stamp\">").Append(rendered).Append("</div>");
            return;
        }

        sb.Append("<div class=\"saveaspdf-stamp\">");
        sb.Append("<h2>SaveAsPDF</h2>");
        sb.Append("<table>");

        if (stamp.IncludeProjectId && !string.IsNullOrWhiteSpace(req?.ProjectId))
            sb.Append("<tr><td class=\"label\">מספר פרויקט:</td><td><b>").Append(Esc(req.ProjectId)).Append("</b></td></tr>");

        if (stamp.IncludeProjectName && !string.IsNullOrWhiteSpace(req?.ProjectName))
            sb.Append("<tr><td class=\"label\">שם פרויקט:</td><td>").Append(Esc(req.ProjectName)).Append("</td></tr>");

        if (stamp.IncludeLeader && !string.IsNullOrWhiteSpace(req?.ProjectLeader))
            sb.Append("<tr><td class=\"label\">מנהל פרויקט:</td><td>").Append(Esc(req.ProjectLeader)).Append("</td></tr>");

        if (stamp.IncludeDate)
            sb.Append("<tr><td class=\"label\">תאריך שמירה:</td><td>")
              .Append(DateTime.Now.ToString("dd/MM/yyyy HH:mm")).Append("</td></tr>");

        // Render every checked field as a row, with a "(אין)" placeholder when
        // the data is missing — so the user can always tell the toggle worked.
        const string Empty = "<i style=\"color:#999\">(אין)</i>";

        if (stamp.IncludeUser)
            sb.Append("<tr><td class=\"label\">נשמר על-ידי:</td><td>")
              .Append(string.IsNullOrWhiteSpace(stamp.UserName) ? Empty : Esc(stamp.UserName))
              .Append("</td></tr>");

        if (stamp.IncludeEmployees)
            sb.Append("<tr><td class=\"label\">עובדי פרויקט:</td><td>")
              .Append(req?.Employees?.Count > 0 ? FormatEmployees(req.Employees) : Empty)
              .Append("</td></tr>");

        if (stamp.IncludeAttachments)
            sb.Append("<tr><td class=\"label\">קבצים מצורפים:</td><td>")
              .Append(stamp.AttachmentNames?.Count > 0 ? FormatAttachmentNames(stamp.AttachmentNames) : Empty)
              .Append("</td></tr>");

        if (stamp.Forwarded)
        {
            sb.Append("<tr><td class=\"label\">הועבר למנהל:</td><td class=\"fwd\">כן");
            if (!string.IsNullOrWhiteSpace(stamp.ForwardedTo))
                sb.Append(" (").Append(Esc(stamp.ForwardedTo)).Append(")");
            sb.Append("</td></tr>");
        }

        if (!string.IsNullOrWhiteSpace(stamp.Notes))
            sb.Append("<tr><td class=\"label\">הערות:</td><td>").Append(Esc(stamp.Notes)).Append("</td></tr>");

        sb.Append("</table>");

        if (stamp.PolicyApplied)
            sb.Append("<div style=\"margin-top:8px;font-size:10px;color:#888;font-style:italic\">[נעול] חלק מהשדות נקבעו על-ידי מנהל המערכת</div>");

        sb.Append("</div>");
    }

    private static void AppendOriginalEmail(StringBuilder sb, EmailDto? email)
    {
        if (email == null) return;
        sb.Append("<div class=\"original\">");
        sb.Append("<h3>").Append(Esc(email.Subject ?? "(ללא נושא)")).Append("</h3>");

        if (!string.IsNullOrWhiteSpace(email.From))
            sb.Append("<div class=\"meta\">מאת: ").Append(Esc(email.From)).Append("</div>");

        if (email.To?.Count > 0)
            sb.Append("<div class=\"meta\">אל: ").Append(Esc(string.Join(", ", email.To))).Append("</div>");

        if (email.Cc?.Count > 0)
            sb.Append("<div class=\"meta\">עותק: ").Append(Esc(string.Join(", ", email.Cc))).Append("</div>");

        if (email.ReceivedDate.HasValue)
            sb.Append("<div class=\"meta\">תאריך: ")
              .Append(email.ReceivedDate.Value.ToString("dd/MM/yyyy HH:mm")).Append("</div>");

        sb.Append("<div style=\"margin-top:10px\">").Append(email.BodyHtml ?? "").Append("</div>");
        sb.Append("</div>");
    }

    // ---------------------------------------------------------------
    // Edge / Chrome headless conversion
    // ---------------------------------------------------------------
    private static (bool ok, string? reason) HtmlToPdf(string htmlPath, string pdfPath)
    {
        var browser = BrowserCandidates.FirstOrDefault(File.Exists);
        if (browser == null)
            return (false, "No headless-capable browser found (Edge/Chrome).");

        var fileUri = "file:///" + htmlPath.Replace('\\', '/');

        var psi = new ProcessStartInfo
        {
            FileName               = browser,
            UseShellExecute        = false,
            CreateNoWindow         = true,
            RedirectStandardOutput = true,
            RedirectStandardError  = true
        };
        psi.ArgumentList.Add("--headless=new");
        psi.ArgumentList.Add("--disable-gpu");
        psi.ArgumentList.Add("--no-pdf-header-footer");
        psi.ArgumentList.Add($"--print-to-pdf={pdfPath}");
        psi.ArgumentList.Add(fileUri);

        try
        {
            using var p = Process.Start(psi);
            if (p == null) return (false, $"Failed to start {browser}.");

            var stderr = p.StandardError.ReadToEnd();
            if (!p.WaitForExit(30000))
            {
                try { p.Kill(); } catch { }
                return (false, "Browser timed out after 30 seconds.");
            }
            if (File.Exists(pdfPath)) return (true, null);

            var trimmed = (stderr ?? "").Trim();
            if (trimmed.Length > 400) trimmed = trimmed[..400] + "…";
            return (false, $"Browser exited {p.ExitCode} without producing PDF. {trimmed}");
        }
        catch (Exception e)
        {
            return (false, e.Message);
        }
    }

    // ---------------------------------------------------------------
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
        var parts = employees.Select(e =>
        {
            var name = e.DisplayName ?? e.Email ?? "";
            var label = Esc(name);
            return e.IsLeader ? $"<b>{label}</b> <span style=\"color:#107c10;font-size:11px\">(מנהל)</span>"
                              : label;
        });
        return string.Join(", ", parts);
    }

    private static string FormatAttachmentNames(List<string>? names)
    {
        if (names == null || names.Count == 0) return "";
        return string.Join(", ", names
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => Esc(n)));
    }
}
