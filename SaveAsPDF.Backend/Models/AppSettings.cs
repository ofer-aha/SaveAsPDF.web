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
        return changed;
    }
}
