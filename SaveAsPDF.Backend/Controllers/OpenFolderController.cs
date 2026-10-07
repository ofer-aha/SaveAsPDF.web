using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

[ApiController]
[Route("api/open-folder")]
public class OpenFolderController : ControllerBase
{
    private readonly SettingsService _settings;
    public OpenFolderController(SettingsService settings) => _settings = settings;

    public class OpenFolderRequest { public string? Path { get; set; } }

    [HttpPost]
    public IActionResult Open([FromBody] OpenFolderRequest req)
    {
        if (string.IsNullOrWhiteSpace(req?.Path))
            return BadRequest(new { error = "Path is required" });

        // Reject NUL and alternate-data-stream syntax before touching the filesystem.
        if (req.Path.Contains('\0'))
            return BadRequest(new { error = "Invalid path" });

        string target;
        try { target = System.IO.Path.GetFullPath(req.Path.Trim()); }
        catch { return BadRequest(new { error = "Invalid path" }); }

        // Security: must be under the configured ProjectsRoot.
        //
        // This check used to be skipped entirely when ProjectsRoot was unset, which
        // meant an unconfigured server would open ANY directory on the machine. Fail
        // closed instead - no configured root means no browsing.
        var root = _settings.Load().ProjectsRoot;
        if (string.IsNullOrWhiteSpace(root))
            return BadRequest(new { error = "Projects root is not configured. Open /admin to set it." });

        var rootFull = System.IO.Path.GetFullPath(root.TrimEnd('\\', '/'));
        var rel      = System.IO.Path.GetRelativePath(rootFull, target);
        if (rel.StartsWith("..", StringComparison.Ordinal) || System.IO.Path.IsPathRooted(rel))
            return BadRequest(new { error = "Path is outside the projects root" });

        if (!Directory.Exists(target))
            return NotFound(new { error = $"Folder does not exist: {target}" });

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName        = "explorer.exe",
                Arguments       = $"\"{target}\"",
                UseShellExecute = true
            });
            return Ok(new { opened = target });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }
}
