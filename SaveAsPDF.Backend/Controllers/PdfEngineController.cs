using Microsoft.AspNetCore.Mvc;

/// <summary>
/// Admin-protected PDF engine info + update (route is under /api/settings so the
/// existing auth middleware applies). Reports the PuppeteerSharp/Chromium engine
/// in use and lets the admin pull the latest Chromium build.
/// </summary>
[ApiController]
[Route("api/settings/pdf/engine")]
public class PdfEngineController : ControllerBase
{
    private readonly SettingsService _settings;
    public PdfEngineController(SettingsService settings) => _settings = settings;

    // ── GET /api/settings/pdf/engine ──────────────────────────────────────────
    [HttpGet]
    public IActionResult Get()
    {
        var (current, installed, cacheDir) = PdfService.GetEngineState();
        return Ok(new
        {
            engine           = "Chromium (headless) via PuppeteerSharp",
            puppeteerVersion = PdfService.PuppeteerVersion,
            currentBuild     = string.IsNullOrEmpty(current) ? "(default pinned build)" : current,
            configuredBuild  = _settings.Load().ChromiumBuild,
            installedBuilds  = installed,
            cacheDir
        });
    }

    // ── POST /api/settings/pdf/engine/update ──────────────────────────────────
    // Downloads the latest available Chromium build and switches the engine to it.
    [HttpPost("update")]
    public async Task<IActionResult> Update()
    {
        try
        {
            var build = await PdfService.UpdateToLatestAsync();

            var s = _settings.Load();
            s.ChromiumBuild = build;
            _settings.Save(s);

            return Ok(new { updated = true, currentBuild = build });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = "Engine update failed: " + ex.Message });
        }
    }
}
