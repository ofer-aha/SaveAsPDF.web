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
        var policy = _settings.Load().StampPolicy ?? new StampPolicy();
        return Ok(new { stamp = policy });
    }
}
