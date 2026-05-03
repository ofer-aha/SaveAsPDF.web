public class PdfResult
{
    public string? FileName { get; set; }
    public string? FullPath { get; set; }

    // True when the actual .pdf was rendered. False when Edge/Chrome failed
    // and the .html fallback was written instead.
    public bool PdfCreated { get; set; }

    // Diagnostic info populated only when PdfCreated == false.
    public string? FallbackReason { get; set; }
}
