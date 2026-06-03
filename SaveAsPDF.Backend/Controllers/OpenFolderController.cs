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

        // Normalise separators and resolve to absolute path
        var target = System.IO.Path.GetFullPath(req.Path.Trim());

        // Security: must be under the configured ProjectsRoot
        var root = _settings.Load().ProjectsRoot;
        if (!string.IsNullOrWhiteSpace(root))
        {
            var rootFull = System.IO.Path.GetFullPath(root.TrimEnd('\\', '/'));
            var rel      = System.IO.Path.GetRelativePath(rootFull, target);
            if (rel.StartsWith("..") || System.IO.Path.IsPathRooted(rel))
                return BadRequest(new { error = "Path is outside the projects root" });
        }

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
