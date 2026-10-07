using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/admin")]
public class AdminAccountController : ControllerBase
{
    private readonly SettingsService _settings;
    public AdminAccountController(SettingsService settings) => _settings = settings;

    public class ChangePasswordRequest
    {
        public string? Username { get; set; }
        public string? Password { get; set; }
    }

    [HttpGet("me")]
    public IActionResult Me()
    {
        var s = _settings.Load();
        return Ok(new
        {
            username             = s.Admin?.Username ?? AdminAuthService.DefaultUsername,
            usingDefaultPassword = AdminAuthService.IsUsingDefaults(s.Admin)
        });
    }

    [HttpPost("password")]
    public IActionResult ChangePassword([FromBody] ChangePasswordRequest body)
    {
        if (body == null || string.IsNullOrWhiteSpace(body.Password))
            return BadRequest(new { error = "סיסמה חדשה נדרשת" });
        // 4 characters is inside brute-force range even through the per-IP throttle,
        // and this password guards the settings API, the server filesystem browser
        // and Process.Start on the console.
        if (body.Password.Length < 12)
            return BadRequest(new { error = "סיסמה צריכה להכיל לפחות 12 תווים" });

        var s = _settings.Load();
        var username = string.IsNullOrWhiteSpace(body.Username)
            ? (s.Admin?.Username ?? AdminAuthService.DefaultUsername)
            : body.Username.Trim();

        s.Admin = AdminAuthService.Hash(username, body.Password);
        _settings.Save(s);

        return Ok(new { username, usingDefaultPassword = false });
    }
}
