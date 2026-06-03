using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/policy")]
public class PolicyController : ControllerBase
{
    private readonly SettingsService _settings;
    public PolicyController(SettingsService settings) => _settings = settings;

    // Public read-only endpoint consumed by the taskpane on load to know
    // which stamp settings the admin has forced. Contains no secrets.
    [HttpGet]
    public IActionResult Get()
    {
        var s = _settings.Load();
        return Ok(new {
            stamp      = s.StampPolicy      ?? new StampPolicy(),
            attachment = s.AttachmentPolicy ?? new AttachmentPolicy()
        });
    }

    // Check if the given email address belongs to the configured admin group.
    // Used by the taskpane to show/hide the admin link. Controls UI only — not a security gate.
    [HttpGet("is-admin")]
    public IActionResult IsAdmin([FromQuery] string? email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return Ok(new { isAdmin = false });

        var groupName = (_settings.Load().AdminGroup ?? "Domain Admins").Trim();
        if (string.IsNullOrWhiteSpace(groupName))
            return Ok(new { isAdmin = false });

        return Ok(new { isAdmin = AdGroupService.IsUserInGroup(email, groupName) });
    }
}
