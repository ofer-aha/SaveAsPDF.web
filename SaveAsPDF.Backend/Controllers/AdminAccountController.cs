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
        if (body.Password.Length < 4)
            return BadRequest(new { error = "סיסמה צריכה להכיל לפחות 4 תווים" });

        var s = _settings.Load();
        var username = string.IsNullOrWhiteSpace(body.Username)
            ? (s.Admin?.Username ?? AdminAuthService.DefaultUsername)
            : body.Username.Trim();

        s.Admin = AdminAuthService.Hash(username, body.Password);
        _settings.Save(s);

        return Ok(new { username, usingDefaultPassword = false });
    }
}
