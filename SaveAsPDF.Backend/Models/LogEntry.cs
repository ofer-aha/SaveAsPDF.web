public class LogEntry
{
    /// <summary>Short random ID used as a React-style key in the frontend table.</summary>
    public string   Id          { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public DateTime Timestamp   { get; set; } = DateTime.Now;

    /// <summary>success | warning | error</summary>
    public string   Level       { get; set; } = "success";

    /// <summary>Outlook display name of the user who initiated the save.</summary>
    public string   Username    { get; set; } = "";

    /// <summary>Subject line of the saved email.</summary>
    public string   Subject     { get; set; } = "";

    /// <summary>Names of attachments included in the save (not necessarily all message attachments).</summary>
    public string[] Attachments { get; set; } = [];

    public string   ProjectId   { get; set; } = "";
    public string   ProjectName { get; set; } = "";

    /// <summary>Full path of the saved PDF (or HTML fallback). Null on error entries.</summary>
    public string?  SavePath    { get; set; }

    /// <summary>Human-readable error description. Null on success entries.</summary>
    public string?  ErrorDetail { get; set; }
}
