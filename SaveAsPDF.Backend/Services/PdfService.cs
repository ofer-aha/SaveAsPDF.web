using System.Text;
using iText.Html2pdf;
using iText.Html2pdf.Resolver.Font;
using iText.Kernel.Pdf;

public static class PdfService
{
    public static PdfResult GeneratePdf(
        EmailDto? email,
        string projectPath,
        SaveAsPdfRequest? request = null)
    {
        var receivedDate = email?.ReceivedDate;
        var timestamp    = receivedDate?.ToString("yyyyMMdd_HHmm")
                           ?? DateTime.Now.ToString("yyyyMMdd_HHmm");

        var safeSubject = Sanitize(StripEmoji(email?.Subject) ?? "(no subject)");
        var fileName    = $"{timestamp}_{safeSubject}.pdf";
        var pdfPath     = Path.Combine(projectPath, fileName);

        var html   = BuildHtml(email, request);
        var result = new PdfResult { FileName = fileName, FullPath = pdfPath, PdfCreated = false };

        // iText7 cannot write directly to UNC paths — generate locally then copy.
        var localTemp = Path.Combine(Path.GetTempPath(), $"saveaspdf_{Guid.NewGuid():N}.pdf");
        try
        {
            ConvertHtmlToPdf(html, localTemp);
            File.Copy(localTemp, pdfPath, overwrite: true);
            result.PdfCreated = true;
        }
        catch (Exception ex)
        {
            // Build full exception chain so the UI shows the root cause, not just "Unknown PdfException"
            var msg = new System.Text.StringBuilder();
            for (var e = ex; e != null; e = e.InnerException)
                msg.Append(e == ex ? "" : " → ").Append('[').Append(e.GetType().Name).Append("] ").Append(e.Message);
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

    // Lazily-built font provider — built once, reused across requests.
    private static DefaultFontProvider? _fontProvider;
    private static readonly object _fontLock = new();

    private static DefaultFontProvider GetFontProvider()
    {
        if (_fontProvider is not null) return _fontProvider;
        lock (_fontLock)
        {
            if (_fontProvider is not null) return _fontProvider;

            // Start with an empty provider and register only fonts that carry
            // Hebrew glyphs (U+0590-U+05FF). The standard PDF fonts (Helvetica,
            // Times-Roman, Courier) do NOT support Hebrew, so we skip them to
            // prevent iText7 from picking them as a fallback.
            var provider = new DefaultFontProvider(false, false, false);

            var fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

            // Ordered by preference: Arial Unicode covers every block; fallback
            // to Arial, Segoe UI, Tahoma which all carry the Hebrew Unicode block.
            var candidates = new[]
            {
                "ARIALUNI.TTF",                           // Arial Unicode MS — widest coverage
                "arial.ttf",  "arialbd.ttf",  "ariali.ttf",  "arialbi.ttf",
                "segoeui.ttf","segoeuib.ttf", "segoeuii.ttf","segoeuiz.ttf",
                "tahoma.ttf", "tahomabd.ttf",
                "times.ttf",  "timesbd.ttf",  "timesi.ttf",  "timesbi.ttf",
                "calibri.ttf","calibrib.ttf",
                "verdana.ttf","verdanab.ttf",
                "cour.ttf",   "courbd.ttf"
            };

            foreach (var name in candidates)
            {
                var path = Path.Combine(fontsDir, name);
                if (File.Exists(path))
                    provider.AddFont(path);
            }

            _fontProvider = provider;
            return _fontProvider;
        }
    }

    private static void ConvertHtmlToPdf(string html, string outputPath)
    {
        // Strip elements that cause iText7 to fail or stall on a server with no
        // internet: scripts, external stylesheets, and external image src URLs.
        html = System.Text.RegularExpressions.Regex.Replace(
            html, @"<script\b[^>]*>[\s\S]*?</script>", "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        html = System.Text.RegularExpressions.Regex.Replace(
            html, @"<link\b[^>]*\brel\s*=\s*[""']stylesheet[""'][^>]*/?>", "",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        html = System.Text.RegularExpressions.Regex.Replace(
            html, @"(src\s*=\s*"")https?://[^""]*""", "$1\"",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        html = System.Text.RegularExpressions.Regex.Replace(
            html, @"(src\s*=\s*')https?://[^']*'", "$1'",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        var props = new ConverterProperties();
        props.SetFontProvider(GetFontProvider());

        using var fs  = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        using var pdf = new PdfDocument(new PdfWriter(fs));
        HtmlConverter.ConvertToPdf(html, pdf, props);
    }

    // ---------------------------------------------------------------
    // HTML composition
    // ---------------------------------------------------------------
    private static string BuildHtml(EmailDto? email, SaveAsPdfRequest? req)
    {
        var sb = new StringBuilder();
        sb.Append(@"<!DOCTYPE html>
<html lang=""he""><head><meta charset=""UTF-8"" />
<style>
  @page { margin: 1.5cm; }
  body {
    font-family: 'Arial Unicode MS', Arial, 'Segoe UI', Tahoma, Verdana, sans-serif;
    color: #222; line-height: 1.5; font-size: 12px; margin: 0;
  }
  /* direction: rtl makes iText7 put the first <td> on the RIGHT (Hebrew label
     column). Character order within cells is handled by RtlText() pre-reversal
     in C# — CSS BiDi properties are ignored by iText7 at the inline level. */
  .saveaspdf-stamp {
    direction: rtl; text-align: right;
    border: 2px solid #0078d4; background: #f0f6fc; border-radius: 6px;
    padding: 12px 16px; margin-bottom: 18px;
  }
  .saveaspdf-stamp table { direction: rtl; border-collapse: collapse; width: 100%; }
  /* direction: ltr prevents iText7 from mirror-substituting brackets/parens.
     text-align: right keeps the pre-reversed Hebrew text visually on the right. */
  .saveaspdf-stamp td   { direction: ltr; text-align: right; padding: 2px 0 2px 8px; vertical-align: top; font-size: 12px; }
  .saveaspdf-stamp .label { color: #666; white-space: nowrap; width: 1%; }
  .saveaspdf-stamp .fwd   { color: #107c10; font-weight: 600; }
  .original { border-top: 1px solid #ddd; padding-top: 14px; direction: rtl; text-align: right; }
  .original .meta { font-size: 11px; color: #666; margin-bottom: 6px; direction: ltr; text-align: right; }
  .original h3    { margin: 0 0 8px 0; font-size: 14px; direction: ltr; text-align: right; }
  .original .body { direction: rtl; unicode-bidi: embed; margin-top: 10px; }
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
        sb.Append("<table>");

        if (stamp.IncludeProjectId && !string.IsNullOrWhiteSpace(req?.ProjectId))
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("מספר פרויקט:")).Append("</td><td><b>")
              .Append(Esc(req.ProjectId)).Append("</b></td></tr>");

        if (stamp.IncludeProjectName && !string.IsNullOrWhiteSpace(req?.ProjectName))
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("שם פרויקט:")).Append("</td><td>")
              .Append(EscRtl(req.ProjectName)).Append("</td></tr>");

        if (stamp.IncludeLeader)
        {
            var leaderEmployee = req?.Employees?.FirstOrDefault(e => e.IsLeader);
            var leaderDisplay  = leaderEmployee?.DisplayName?.Trim();
            var leaderEmail    = leaderEmployee?.Email?.Trim();
            if (string.IsNullOrWhiteSpace(leaderDisplay))
                leaderDisplay = req?.ProjectLeader?.Trim();
            if (!string.IsNullOrWhiteSpace(leaderDisplay))
            {
                var leaderCell = !string.IsNullOrWhiteSpace(leaderEmail)
                    ? $"<a href=\"mailto:{Esc(leaderEmail)}\">{EscRtl(leaderDisplay)}</a>"
                    : EscRtl(leaderDisplay);
                sb.Append("<tr><td class=\"label\">").Append(EscRtl("מנהל פרויקט:"))
                  .Append("</td><td>").Append(leaderCell).Append("</td></tr>");
            }
        }

        if (stamp.IncludeDate)
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("תאריך שמירה:")).Append("</td><td>")
              .Append(Esc(DateTime.Now.ToString("dd/MM/yyyy HH:mm"))).Append("</td></tr>");

        var empty = $"<i style=\"color:#999\">{EscRtl("(אין)")}</i>";

        if (stamp.IncludeUser)
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("נשמר על-ידי:")).Append("</td><td>")
              .Append(string.IsNullOrWhiteSpace(stamp.UserName) ? empty : EscRtl(stamp.UserName))
              .Append("</td></tr>");

        if (stamp.IncludeEmployees)
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("עובדי פרויקט:")).Append("</td><td>")
              .Append(req?.Employees?.Count > 0 ? FormatEmployees(req.Employees) : empty)
              .Append("</td></tr>");

        if (stamp.IncludeAttachments)
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("קבצים מצורפים:")).Append("</td><td>")
              .Append(stamp.AttachmentNames?.Count > 0 ? FormatAttachmentNames(stamp.AttachmentNames) : empty)
              .Append("</td></tr>");

        if (stamp.Forwarded)
        {
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("הועבר למנהל:"))
              .Append("</td><td class=\"fwd\">").Append(EscRtl("כן"));
            if (!string.IsNullOrWhiteSpace(stamp.ForwardedTo))
                sb.Append(" (").Append(EscRtl(stamp.ForwardedTo)).Append(")");
            sb.Append("</td></tr>");
        }

        if (!string.IsNullOrWhiteSpace(stamp.Notes))
            sb.Append("<tr><td class=\"label\">").Append(EscRtl("הערות:")).Append("</td><td>")
              .Append(EscRtl(stamp.Notes)).Append("</td></tr>");

        sb.Append("</table>");

        if (stamp.PolicyApplied)
            sb.Append("<div style=\"margin-top:8px;font-size:10px;color:#888;font-style:italic\">")
              .Append(EscRtl("[נעול] חלק מהשדות נקבעו על-ידי מנהל המערכת")).Append("</div>");

        sb.Append("</div>");
    }

    private static void AppendOriginalEmail(StringBuilder sb, EmailDto? email)
    {
        if (email == null) return;
        sb.Append("<div class=\"original\">");
        // Subject: may be Hebrew — pre-reverse
        sb.Append("<h3>").Append(EscRtl(StripEmoji(email.Subject) ?? "(ללא נושא)")).Append("</h3>");

        // Meta lines: Hebrew label pre-reversed, LTR value follows after a space
        if (!string.IsNullOrWhiteSpace(email.From))
            sb.Append("<div class=\"meta\">").Append(EscRtl("מאת:")).Append(" ").Append(Esc(email.From)).Append("</div>");

        if (email.To?.Count > 0)
            sb.Append("<div class=\"meta\">").Append(EscRtl("אל:")).Append(" ").Append(Esc(string.Join(", ", email.To))).Append("</div>");

        if (email.Cc?.Count > 0)
            sb.Append("<div class=\"meta\">").Append(EscRtl("עותק:")).Append(" ").Append(Esc(string.Join(", ", email.Cc))).Append("</div>");

        if (email.ReceivedDate.HasValue)
            sb.Append("<div class=\"meta\">").Append(EscRtl("תאריך:")).Append(" ")
              .Append(Esc(email.ReceivedDate.Value.ToString("dd/MM/yyyy HH:mm"))).Append("</div>");

        // Email body: apply run-level RTL pre-processing to every text node so
        // iText7's LTR renderer produces correct Hebrew visual order.
        sb.Append("<div class=\"body\">").Append(ApplyRtlToHtml(email.BodyHtml ?? "")).Append("</div>");
        sb.Append("</div>");
    }

    // ---------------------------------------------------------------
    private static string Sanitize(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }

    // Remove characters outside the BMP (emoji, supplementary symbols) that
    // registered Windows fonts don't have glyphs for. iText7 throws an
    // ArgumentOutOfRangeException when it encounters orphaned surrogates.
    private static string StripEmoji(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var sb = new System.Text.StringBuilder(s.Length);
        for (int i = 0; i < s.Length; )
        {
            if (char.IsHighSurrogate(s[i]) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
            {
                i += 2; // skip the surrogate pair (emoji / supplementary char)
            }
            else
            {
                sb.Append(s[i]);
                i++;
            }
        }
        return sb.ToString();
    }

    private static string Esc(string? s) =>
        string.IsNullOrEmpty(s) ? "" :
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    // iText7 ignores all CSS BiDi properties and renders every glyph LTR.
    // We implement a simplified Unicode BiDi visual reorder so text reads
    // correctly when a Hebrew user scans the page right-to-left:
    //
    //  1. Parse the string into Unicode code points (keeps surrogate pairs intact).
    //  2. Classify each code point: Hebrew (H), Latin/digit (L), or neutral (N).
    //  3. Resolve neutrals: they inherit the direction of the preceding strong char.
    //  4. Merge adjacent same-direction code points into runs.
    //  5. Reverse the run order (RTL base direction).
    //  6. Within each Hebrew run, also reverse the character order.
    //     LTR runs (e.g. "Re:", email addresses, numbers) stay intact.
    //
    // Result: "Re: כדי לשפר" → visual "רפשל ידכ Re:" — Hebrew reads correctly
    // right-to-left; Latin words keep their natural left-to-right letter order.
    private static string RtlText(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";

        // --- Step 1: parse into code points ---
        var cp    = new List<string>(s.Length);
        var cpDir = new List<char>(s.Length); // 'H', 'L', or 'N'
        bool hasHebrew = false;

        for (int i = 0; i < s.Length; )
        {
            string unit;
            char   dir;
            if (char.IsHighSurrogate(s[i]) && i + 1 < s.Length && char.IsLowSurrogate(s[i + 1]))
            {
                unit = new string(new[] { s[i], s[i + 1] });
                dir  = 'N'; // emoji / supplementary = neutral
                i += 2;
            }
            else
            {
                char c = s[i++];
                unit = c.ToString();
                if (c is >= 'א' and <= 'ת')                          { dir = 'H'; hasHebrew = true; }
                else if (c is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z') or (>= '0' and <= '9'))
                                                                      { dir = 'L'; }
                else                                                  { dir = 'N'; }
            }
            cp.Add(unit);
            cpDir.Add(dir);
        }

        if (!hasHebrew) return s;

        // --- Step 2: resolve neutrals (inherit preceding strong direction) ---
        char last = 'L';
        for (int i = 0; i < cpDir.Count; i++)
        {
            if (cpDir[i] != 'N') { last = cpDir[i]; }
            else                 { cpDir[i] = last;  }
        }

        // --- Step 3: group into runs ---
        var runs = new List<(char dir, List<string> chars)>();
        for (int i = 0; i < cp.Count; i++)
        {
            if (runs.Count == 0 || runs[^1].dir != cpDir[i])
                runs.Add((cpDir[i], new List<string>()));
            runs[^1].chars.Add(cp[i]);
        }

        // --- Step 4: reverse run order; reverse chars inside Hebrew runs ---
        runs.Reverse();
        var sb = new StringBuilder();
        foreach (var (dir, chars) in runs)
        {
            if (dir == 'H') for (int i = chars.Count - 1; i >= 0; i--) sb.Append(chars[i]);
            else            foreach (var c in chars)                     sb.Append(c);
        }
        return sb.ToString();
    }

    // Apply RtlText to every plain-text node inside an HTML fragment.
    // Tags and attributes are passed through unchanged; only visible text is reordered.
    private static string ApplyRtlToHtml(string html)
    {
        if (string.IsNullOrEmpty(html)) return "";
        var result = new StringBuilder(html.Length);
        int pos = 0;
        while (pos < html.Length)
        {
            // Find next tag opening
            int tagOpen = html.IndexOf('<', pos);
            if (tagOpen < 0) tagOpen = html.Length;

            // Text node before the tag
            if (tagOpen > pos)
            {
                var raw     = html.Substring(pos, tagOpen - pos);
                var decoded = System.Net.WebUtility.HtmlDecode(raw);
                result.Append(Esc(RtlText(decoded)));
            }

            if (tagOpen >= html.Length) break;

            // Copy the tag itself verbatim (preserve all attributes including dir="rtl")
            int tagClose = html.IndexOf('>', tagOpen);
            if (tagClose < 0) { result.Append(html.Substring(tagOpen)); break; }

            // Skip script/style blocks entirely
            var tag = html.Substring(tagOpen, tagClose - tagOpen + 1).ToLowerInvariant();
            if (tag.StartsWith("<script") || tag.StartsWith("<style"))
            {
                string close = tag.StartsWith("<script") ? "</script>" : "</style>";
                int blockEnd = html.IndexOf(close, tagClose, StringComparison.OrdinalIgnoreCase);
                if (blockEnd < 0) { result.Append(html.Substring(tagOpen)); break; }
                pos = blockEnd + close.Length;
                continue;
            }

            result.Append(html.Substring(tagOpen, tagClose - tagOpen + 1));
            pos = tagClose + 1;
        }
        return result.ToString();
    }

    // HTML-escape after RTL pre-processing so entity characters (&, <, >)
    // are never split by the reversal.
    private static string EscRtl(string? s) => Esc(RtlText(s));

    private static string FormatEmployees(List<EmployeeDto>? employees)
    {
        if (employees == null || employees.Count == 0) return "";
        var parts = employees.Select(e =>
        {
            var name  = e.DisplayName ?? e.Email ?? "";
            var label = EscRtl(name);
            return e.IsLeader
                ? $"<b>{label}</b> <span style=\"color:#107c10;font-size:11px\">{EscRtl("(מנהל)")}</span>"
                : label;
        });
        return string.Join(EscRtl(" ,"), parts);  // comma+space also reversed for RTL
    }

    private static string FormatAttachmentNames(List<string>? names)
    {
        if (names == null || names.Count == 0) return "";
        return string.Join(", ", names
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => EscRtl(n)));
    }
}
