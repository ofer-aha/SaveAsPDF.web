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

    // -- POST /api/settings/pdf/engine/select ---------------------------------
    // Switches the engine to an already-installed Chromium build (or back to
    // PuppeteerSharp's default pinned build when build is empty/null). This is the
    // rollback path for an engine update that turns out to be incompatible with the
    // bundled PuppeteerSharp version.
    public class SelectBuildRequest { public string? Build { get; set; } }

    [HttpPost("select")]
    public async Task<IActionResult> Select([FromBody] SelectBuildRequest? body)
    {
        try
        {
            var build = await PdfService.SelectBuildAsync(body?.Build);

            var s = _settings.Load();
            s.ChromiumBuild = string.IsNullOrEmpty(build) ? null : build;
            _settings.Save(s);

            return Ok(new
            {
                selected     = true,
                currentBuild = string.IsNullOrEmpty(build) ? "(default pinned build)" : build
            });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = "Engine switch failed: " + ex.Message });
        }
    }
}
