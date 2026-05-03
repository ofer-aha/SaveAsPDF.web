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
            projectsRoot = s.ProjectsRoot,
            stampPolicy  = s.StampPolicy
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

        _settings.Save(new AppSettings
        {
            ProjectsRoot = root,
            StampPolicy  = incoming.StampPolicy ?? new StampPolicy()
        });

        var saved = _settings.Load();
        return Ok(new
        {
            projectsRoot = saved.ProjectsRoot,
            stampPolicy  = saved.StampPolicy,
            warning
        });
    }
}
