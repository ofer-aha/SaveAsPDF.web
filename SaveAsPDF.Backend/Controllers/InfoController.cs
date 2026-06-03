using Microsoft.AspNetCore.Mvc;
using System.Reflection;
using System.Text.Json;

[ApiController]
[Route("api/info")]
public class InfoController : ControllerBase
{
    [HttpGet]
    public IActionResult Get()
    {
        // Read version from package.json deployed alongside the binary.
        // This is the single source of truth — no rebuild needed when bumping the version.
        var version = ReadPackageJsonVersion()
                   ?? typeof(InfoController).Assembly
                          .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                          ?.InformationalVersion?.Split('+')[0]
                   ?? "0.0.0";

        return Ok(new
        {
            name    = "SaveAsPDF.Backend",
            version,
            author  = "Ofer Aharon",
            email   = "ofer@sw-eng.co.il"
        });
    }

    private static string? ReadPackageJsonVersion()
    {
        try
        {
            var path = Path.Combine(AppContext.BaseDirectory, "package.json");
            if (!System.IO.File.Exists(path)) return null;
            using var doc = JsonDocument.Parse(System.IO.File.ReadAllText(path));
            return doc.RootElement.GetProperty("version").GetString();
        }
        catch { return null; }
    }
}
