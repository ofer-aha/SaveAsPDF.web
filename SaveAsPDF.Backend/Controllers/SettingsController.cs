using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/settings")]
public class SettingsController : ControllerBase
{
    private readonly SettingsService _settings;
    public SettingsController(SettingsService settings) => _settings = settings;

    [HttpGet]
    public IActionResult Get()
    {
        // Never expose the password hash/salt over the wire — even to authenticated
        // admins. Use /api/admin/me to read account state.
        var s = _settings.Load();
        return Ok(new
        {
            projectsRoot     = s.ProjectsRoot,
            adminGroup       = s.AdminGroup,
            stampPolicy      = s.StampPolicy,
            attachmentPolicy = s.AttachmentPolicy
        });
    }

    [HttpPost]
    public IActionResult Save([FromBody] AppSettings incoming)
    {
        if (incoming == null)
            return BadRequest(new { error = "Body required" });

        var root = (incoming.ProjectsRoot ?? "").Trim();
        if (string.IsNullOrEmpty(root))
            return BadRequest(new { error = "Projects root folder cannot be empty" });

        // Warn but don't block: the folder may be unmounted/offline right now
        // (e.g. network drive). The policy edits should still persist; a real
        // SaveAsPDF call will surface the missing-folder error at use time.
        string? warning = Directory.Exists(root)
            ? null
            : $"Warning: folder does not exist or is offline: {root}";

        // Load current settings so admin credentials and other sections are
        // preserved. Without this, every settings save would reset them.
        var current = _settings.Load();

        // Assign only what this form actually owns. Rebuilding the object from
        // scratch meant any field the General tab does not post was reset to its
        // default on every save: ChromiumBuild was not in the initializer at all,
        // so choosing a Chromium build on the PDF tab and then fixing a typo here
        // silently reverted the engine; AdminGroup and AttachmentPolicy have no UI
        // on this page either, yet "?? default" overwrote whatever was configured.
        current.ProjectsRoot = root;
        if (incoming.AdminGroup       != null) current.AdminGroup       = incoming.AdminGroup.Trim();
        if (incoming.StampPolicy      != null) current.StampPolicy      = incoming.StampPolicy;
        if (incoming.AttachmentPolicy != null) current.AttachmentPolicy = incoming.AttachmentPolicy;

        _settings.Save(current);

        var saved = _settings.Load();
        return Ok(new
        {
            projectsRoot     = saved.ProjectsRoot,
            adminGroup       = saved.AdminGroup,
            stampPolicy      = saved.StampPolicy,
            attachmentPolicy = saved.AttachmentPolicy,
            warning
        });
    }

    // ── GET /api/settings/browse ──────────────────────────────────────────────
    // Server-side folder picker for the admin "Browse" button next to Projects
    // Root Folder. Admin-protected (route is under /api/settings). With no path,
    // returns the machine's drive roots; otherwise the subfolders of `path`.
    [HttpGet("browse")]
    public IActionResult Browse([FromQuery] string? path)
    {
        try
        {
            // No path → list drive roots (e.g. C:\, J:\)
            if (string.IsNullOrWhiteSpace(path))
            {
                var drives = DriveInfo.GetDrives()
                    .Where(d => d.IsReady || d.DriveType == DriveType.Network)
                    .Select(d => new { name = d.Name, path = d.Name })
                    .ToList();
                return Ok(new { current = (string?)null, parent = (string?)null, dirs = drives });
            }

            var full = Path.GetFullPath(path);
            if (!Directory.Exists(full))
                return NotFound(new { error = "Folder does not exist: " + full });

            var parent = Path.GetDirectoryName(full.TrimEnd(Path.DirectorySeparatorChar));

            var names = new List<string>();
            try { names.AddRange(Directory.EnumerateDirectories(full).Select(Path.GetFileName)!); }
            catch (UnauthorizedAccessException) { /* skip unreadable subtrees */ }

            var dirs = names
                .Where(n => !string.IsNullOrEmpty(n) && !n!.StartsWith("."))
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .Select(n => new { name = n, path = Path.Combine(full, n!) })
                .ToList();

            return Ok(new { current = full, parent, dirs });
        }
        catch (Exception ex)
        {
            return BadRequest(new { error = ex.Message });
        }
    }

    // ── GET /api/settings/pdf ─────────────────────────────────────────────────

    [HttpGet("pdf")]
    public IActionResult GetPdf()
    {
        var s = _settings.Load();
        return Ok(new { settings = s.PdfSettings, policy = s.PdfPolicy ?? new PdfPolicy() });
    }

    // ── POST /api/settings/pdf ────────────────────────────────────────────────
    // Body: { settings: PdfSettings, policy: PdfPolicy }. The policy records which
    // fields the admin controls (locked) in the user taskpane.

    [HttpPost("pdf")]
    public IActionResult SavePdf([FromBody] PdfSettingsSaveRequest? incoming)
    {
        var settingsIn = incoming?.Settings;
        if (settingsIn == null)
            return BadRequest(new { error = "Body required" });

        // Clamp margins to a sane range (0.5 – 10 cm)
        static double Clamp(double v) => Math.Max(0.5, Math.Min(10.0, v));

        var validated = new PdfSettings
        {
            PageSize        = settingsIn.PageSize is "A4" or "Letter" or "Legal" or "A3"
                                ? settingsIn.PageSize : "A4",
            Landscape       = settingsIn.Landscape,
            MarginTopCm     = Clamp(settingsIn.MarginTopCm),
            MarginBottomCm  = Clamp(settingsIn.MarginBottomCm),
            MarginLeftCm    = Clamp(settingsIn.MarginLeftCm),
            MarginRightCm   = Clamp(settingsIn.MarginRightCm),
            PrintBackground = settingsIn.PrintBackground,

            // Render controls. Clamped so a bad value cannot wedge the renderer:
            // too low and nothing ever finishes, too high and a stuck page blocks
            // a request for minutes.
            RenderTimeoutMs    = Math.Clamp(settingsIn.RenderTimeoutMs <= 0
                                                ? 15000 : settingsIn.RenderTimeoutMs,
                                            2000, 120000),
            AllowRemoteContent = settingsIn.AllowRemoteContent,
            AllowRemoteImages  = settingsIn.AllowRemoteImages,

            // Same reasoning as the render timeout: a per-image deadline of 50 ms
            // fetches nothing, and one of two minutes hands a mail full of dead
            // image URLs the power to hold a save open.
            RemoteImageTimeoutMs = Math.Clamp(settingsIn.RemoteImageTimeoutMs <= 0
                                                  ? 4000 : settingsIn.RemoteImageTimeoutMs,
                                              1000, 15000),
            MergePdfAttachments  = settingsIn.MergePdfAttachments
        };

        var current = _settings.Load();
        current.PdfSettings = validated;
        current.PdfPolicy   = incoming?.Policy ?? new PdfPolicy();
        _settings.Save(current);

        return Ok(new { settings = validated, policy = current.PdfPolicy });
    }

    public class PdfSettingsSaveRequest
    {
        public PdfSettings? Settings { get; set; }
        public PdfPolicy?   Policy   { get; set; }
    }
}
