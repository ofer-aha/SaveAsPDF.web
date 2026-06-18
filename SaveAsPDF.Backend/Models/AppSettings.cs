public class AppSettings
{
    public string ProjectsRoot { get; set; } = @"J:\";

    // AD group whose members see the admin link in the taskpane header.
    public string AdminGroup { get; set; } = "Domain Admins";

    // Admin-forced stamp settings. Each nullable bool: null = user controls,
    // true/false = forced value (taskpane locks the toggle, backend overrides on save).
    public StampPolicy StampPolicy { get; set; } = new();

    // Admin-forced attachment settings.
    public AttachmentPolicy AttachmentPolicy { get; set; } = new();

    // Admin login credentials. When null/empty PasswordHash, defaults to admin/admin
    // (the auth service falls back to this so first-time setup is possible).
    public AdminCredentials? Admin { get; set; }

    // Log retention settings — how many entries to keep and for how long.
    public LogSettings LogSettings { get; set; } = new();

    // PDF output settings — page size, margins, orientation.
    // These are the admin defaults; users may override unlocked fields from the
    // taskpane (see PdfPolicy).
    public PdfSettings PdfSettings { get; set; } = new();

    // Which PDF settings the admin controls. A locked field forces the admin's
    // PdfSettings value and disables the matching control in the user taskpane.
    public PdfPolicy PdfPolicy { get; set; } = new();

    // Selected Chromium build for the PDF engine (PuppeteerSharp BrowserFetcher
    // buildId). null/empty = use PuppeteerSharp's default pinned build. Set by the
    // admin "update engine" action and applied on startup.
    public string? ChromiumBuild { get; set; }
}

public class PdfPolicy
{
    // true = admin-controlled (locked to the admin's PdfSettings value);
    // false = user controls the field from the taskpane.
    public bool PageSize        { get; set; }
    public bool Orientation     { get; set; }
    public bool Margins         { get; set; }
    public bool PrintBackground { get; set; }

    // Build the effective PdfSettings from the admin defaults + a user's chosen
    // values. For each field: locked → admin value; otherwise the user value
    // (falling back to admin when the user sent nothing).
    public PdfSettings Resolve(PdfSettings admin, PdfSettings? user)
    {
        static double Clamp(double v) => Math.Max(0.5, Math.Min(10.0, v));
        var u = user ?? admin;
        return new PdfSettings
        {
            PageSize        = PageSize        ? admin.PageSize        : u.PageSize,
            Landscape       = Orientation     ? admin.Landscape       : u.Landscape,
            MarginTopCm     = Margins         ? admin.MarginTopCm     : Clamp(u.MarginTopCm),
            MarginBottomCm  = Margins         ? admin.MarginBottomCm  : Clamp(u.MarginBottomCm),
            MarginLeftCm    = Margins         ? admin.MarginLeftCm    : Clamp(u.MarginLeftCm),
            MarginRightCm   = Margins         ? admin.MarginRightCm   : Clamp(u.MarginRightCm),
            PrintBackground = PrintBackground ? admin.PrintBackground : u.PrintBackground
        };
    }
}

public class PdfSettings
{
    // Page size string: "A4" | "Letter" | "Legal" | "A3"
    public string PageSize        { get; set; } = "A4";

    // Page orientation
    public bool   Landscape       { get; set; } = false;

    // Margins in centimetres — Word 2016+ default is 2.54 cm (1 inch) on all sides.
    public double MarginTopCm     { get; set; } = 2.54;
    public double MarginBottomCm  { get; set; } = 2.54;
    public double MarginLeftCm    { get; set; } = 2.54;
    public double MarginRightCm   { get; set; } = 2.54;

    // Render CSS backgrounds (colours, images). Keep on by default for the stamp banner.
    public bool   PrintBackground { get; set; } = true;
}

public class LogSettings
{
    // Maximum number of log entries kept in the in-memory ring buffer.
    public int MaxEntries { get; set; } = 1000;
    // Auto-purge entries older than this many days. 0 = keep forever.
    public int RetainDays { get; set; } = 30;
}

public class AttachmentPolicy
{
    // Signature-image filter threshold in bytes.
    // null  = user controls their own setting
    //   0   = admin forces the feature OFF (show all attachments)
    // > 0   = admin forces the feature ON with this threshold
    public int? SigImgThreshold { get; set; }
}

public class AdminCredentials
{
    public string Username     { get; set; } = "";
    public string PasswordHash { get; set; } = "";
    public string PasswordSalt { get; set; } = "";
}

public class StampPolicy
{
    public bool? DefaultStamp        { get; set; }
    public bool? IncludeProjectId    { get; set; }
    public bool? IncludeProjectName  { get; set; }
    public bool? IncludeLeader       { get; set; }
    public bool? IncludeDate         { get; set; }
    public bool? IncludeUser         { get; set; }
    public bool? IncludeEmployees    { get; set; }
    public bool? IncludeAttachments  { get; set; }

    // Original-message header fields.
    public bool? IncludeFrom         { get; set; }
    public bool? IncludeTo           { get; set; }
    public bool? IncludeCc           { get; set; }
    public bool? IncludeSent         { get; set; }
    public bool? IncludeReceived     { get; set; }
    public bool? IncludeSubject      { get; set; }

    // Apply the policy onto an incoming StampInfo (mutates in place).
    // Returns true if any field was overridden.
    public bool ApplyTo(StampInfo info)
    {
        var changed = false;
        if (IncludeProjectId    is bool a) { info.IncludeProjectId    = a; changed = true; }
        if (IncludeProjectName  is bool b) { info.IncludeProjectName  = b; changed = true; }
        if (IncludeLeader       is bool c) { info.IncludeLeader       = c; changed = true; }
        if (IncludeDate         is bool d) { info.IncludeDate         = d; changed = true; }
        if (IncludeUser         is bool e) { info.IncludeUser         = e; changed = true; }
        if (IncludeEmployees    is bool f) { info.IncludeEmployees    = f; changed = true; }
        if (IncludeAttachments  is bool g) { info.IncludeAttachments  = g; changed = true; }
        if (IncludeFrom         is bool h) { info.IncludeFrom         = h; changed = true; }
        if (IncludeTo           is bool i) { info.IncludeTo           = i; changed = true; }
        if (IncludeCc           is bool j) { info.IncludeCc           = j; changed = true; }
        if (IncludeSent         is bool k) { info.IncludeSent         = k; changed = true; }
        if (IncludeReceived     is bool l) { info.IncludeReceived     = l; changed = true; }
        if (IncludeSubject      is bool m) { info.IncludeSubject      = m; changed = true; }
        return changed;
    }
}
