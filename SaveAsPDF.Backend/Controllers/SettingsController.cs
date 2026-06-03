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

        // Load current settings so admin credentials (password hash/salt) are
        // preserved. Without this, every settings save would reset Admin to null
        // and the auth service would fall back to the default "admin/admin" password.
        var current = _settings.Load();

        _settings.Save(new AppSettings
        {
            ProjectsRoot     = root,
            AdminGroup       = (incoming.AdminGroup ?? "Domain Admins").Trim(),
            StampPolicy      = incoming.StampPolicy      ?? new StampPolicy(),
            AttachmentPolicy = incoming.AttachmentPolicy ?? new AttachmentPolicy(),
            Admin            = current.Admin   // preserve existing credentials
        });

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
}
