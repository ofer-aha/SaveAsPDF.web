using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

/// <summary>
/// Appends the PDF-type attachments of a saved message to the end of the PDF that
/// was generated for it, so one file holds the mail and everything that came with
/// it. Attachments are still written next to the PDF as separate files - this adds
/// a combined copy, it does not replace them.
///
/// Best effort by design: a password-protected, damaged or exotic attachment is
/// skipped and the original PDF is left exactly as it was. Nothing here may cost
/// the user their save.
/// </summary>
public static class PdfMergeService
{
    // Ceilings, not expectations. A 900-page scan attached to a one-line e-mail is
    // a mistake, and merging it would make the archived file unusable.
    private const int  MaxTotalPages       = 500;
    private const long MaxAttachmentBytes  = 40L * 1024 * 1024;

    /// <summary>
    /// Returns the number of attachments appended. 0 means the PDF was not touched.
    /// </summary>
    public static int AppendPdfAttachments(string? pdfPath, List<AttachmentDto>? attachments)
    {
        if (string.IsNullOrWhiteSpace(pdfPath) || !File.Exists(pdfPath)) return 0;

        var candidates = (attachments ?? [])
            .Where(a => !string.IsNullOrEmpty(a?.Base64) && LooksLikePdf(a))
            .ToList();
        if (candidates.Count == 0) return 0;

        PdfDocument? target = null;
        var appended = 0;

        try
        {
            // Open from a copy in memory rather than from the path. PdfReader holds
            // the file open until Close(), and saving back over a path it is still
            // reading from is how you end up with a truncated PDF and no original.
            var original = File.ReadAllBytes(pdfPath);
            using var baseStream = new MemoryStream(original, writable: false);
            target = PdfReader.Open(baseStream, PdfDocumentOpenMode.Modify);

            foreach (var att in candidates)
            {
                try
                {
                    var clean = att.Base64!.Replace("\r", "").Replace("\n", "")
                                           .Replace("\t", "").Replace(" ", "");
                    var bytes = Convert.FromBase64String(clean);
                    if (bytes.LongLength is 0 or > MaxAttachmentBytes) continue;

                    using var source = new MemoryStream(bytes, writable: false);
                    // Import mode copies pages without re-parsing content streams,
                    // which is both faster and far more tolerant of odd producers
                    // (copier firmware, in particular) than a full open.
                    using var doc = PdfReader.Open(source, PdfDocumentOpenMode.Import);

                    if (target.PageCount + doc.PageCount > MaxTotalPages)
                    {
                        Console.Error.WriteLine(
                            $"[SaveAsPDF] merge page budget reached; '{att.Name}' not appended.");
                        break;
                    }

                    for (var i = 0; i < doc.PageCount; i++)
                        target.AddPage(doc.Pages[i]);

                    appended++;
                }
                catch (Exception ex)
                {
                    // Encrypted, truncated, or simply not a PDF despite its name.
                    // The standalone file next to the PDF is still there.
                    Console.Error.WriteLine(
                        $"[SaveAsPDF] could not append '{att.Name}' to the PDF: {ex.Message}");
                }
            }

            if (appended > 0) target.Save(pdfPath);
            return appended;
        }
        catch (Exception ex)
        {
            // Could not even open our own output - leave it untouched and move on.
            Console.Error.WriteLine($"[SaveAsPDF] PDF merge failed: {ex.Message}");
            return 0;
        }
        finally
        {
            try { target?.Close(); } catch { }
        }
    }

    // Trust the bytes over the name: an attachment called .pdf that is really a
    // Word file must not be fed to the parser, and a PDF saved without an
    // extension should still merge.
    private static bool LooksLikePdf(AttachmentDto att)
    {
        var byName = (att.Name ?? "").EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
        var byType = (att.ContentType ?? "").Contains("pdf", StringComparison.OrdinalIgnoreCase);
        if (!byName && !byType) return false;

        // "%PDF" is the first four bytes of every PDF; base64 encodes them as "JVBER".
        var head = att.Base64!.TrimStart();
        return head.StartsWith("JVBER", StringComparison.Ordinal);
    }
}
