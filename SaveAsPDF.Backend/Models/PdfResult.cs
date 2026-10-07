public class PdfResult
{
    public string? FileName { get; set; }
    public string? FullPath { get; set; }

    // True when the actual .pdf was rendered. False when Edge/Chrome failed
    // and the .html fallback was written instead.
    public bool PdfCreated { get; set; }

    // Diagnostic info populated only when PdfCreated == false.
    public string? FallbackReason { get; set; }

    // True when SOMETHING reached disk - the PDF, or the HTML fallback. False
    // means the save produced no file at all (share offline, no permission, disk
    // full, path too long). Callers must not report success on false: the caller
    // used to mark the e-mail as archived on the strength of FullPath alone, which
    // is set even when every write threw.
    public bool FileWritten { get; set; }
}
