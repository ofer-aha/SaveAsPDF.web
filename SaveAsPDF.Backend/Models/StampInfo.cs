public class StampInfo
{
    // Selected fields to include
    public bool IncludeProjectId   { get; set; } = true;
    public bool IncludeProjectName { get; set; } = true;
    public bool IncludeLeader      { get; set; } = true;
    public bool IncludeDate        { get; set; } = true;
    public bool IncludeUser        { get; set; }
    public bool IncludeEmployees   { get; set; }
    public bool IncludeAttachments { get; set; }

    // Original-message header fields rendered inside the SaveAsPDF frame.
    public bool IncludeFrom     { get; set; } = true;
    public bool IncludeTo       { get; set; } = true;
    public bool IncludeCc       { get; set; } = true;
    public bool IncludeSent     { get; set; } = true;
    public bool IncludeReceived { get; set; } = true;
    public bool IncludeSubject  { get; set; } = true;

    public string? Notes    { get; set; }
    public string? Template { get; set; }   // optional override; non-empty = use as HTML

    // Set true when the dispatcher chose to forward to the project leader
    public bool   Forwarded   { get; set; }
    public string? ForwardedTo { get; set; }

    // Display name of the Outlook user who initiated the save
    public string? UserName { get; set; }

    // All attachment names on the original message — used only for stamping
    // (independent of which attachments the user chose to save).
    public List<string>? AttachmentNames { get; set; }

    // Set server-side when admin policy overrode any field of this stamp.
    // Renders a small footer note inside the PDF stamp for transparency.
    public bool PolicyApplied { get; set; }
}
