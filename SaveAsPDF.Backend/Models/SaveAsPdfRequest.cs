public class SaveAsPdfRequest
{
	public string? ProjectId      { get; set; }
	public string? ProjectName    { get; set; }
	public string? ProjectLeader  { get; set; }

	// Display name of the Outlook user who clicked Save (sent regardless of stamp settings).
	public string? SavedBy { get; set; }

	public EmailDto? Email                  { get; set; }
	public List<AttachmentDto>? Attachments { get; set; }
	public List<EmployeeDto>? Employees     { get; set; }

	// Optional sub-folder under the resolved project root, e.g. "Drafts" or "2026/Phase 1".
	// Forward slashes; empty / null = save in project root.
	public string? DestinationFolder { get; set; }

	// Stamp / forward information embedded into the PDF
	public StampInfo? Stamp { get; set; }

	// User-chosen PDF output settings (page size, orientation, margins, background).
	// The server merges these with the admin PdfPolicy: locked fields are forced
	// to the admin value, unlocked fields use what the user sent here.
	public PdfSettings? PdfSettings { get; set; }
}
