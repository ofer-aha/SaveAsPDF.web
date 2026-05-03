using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/project")]
public class ProjectController : ControllerBase
{
    private readonly SettingsService _settings;
    private readonly ProjectDataService _data;

    public ProjectController(SettingsService settings, ProjectDataService data)
    {
        _settings = settings;
        _data     = data;
    }

    // Create the project's root folder if it doesn't exist.
    // Used by the taskpane after the user confirms "create new project" on a missing number.
    [HttpPost("{number}")]
    public IActionResult Create(string number)
    {
        if (string.IsNullOrWhiteSpace(number))
            return BadRequest(new { error = "Project number is required" });

        var settings = _settings.Load();
        if (string.IsNullOrWhiteSpace(settings.ProjectsRoot))
            return BadRequest(new { error = "Projects root folder is not configured" });

        DirectoryInfo dir;
        try { dir = ProjectPathResolver.Resolve(number, settings.ProjectsRoot); }
        catch (Exception ex) { return BadRequest(new { error = ex.Message }); }

        try { Directory.CreateDirectory(dir.FullName); }
        catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }

        return Ok(new { folderPath = dir.FullName, projectNumber = number, created = true });
    }

    [HttpGet("{number}")]
    public IActionResult Get(string number)
    {
        if (string.IsNullOrWhiteSpace(number))
            return BadRequest(new { error = "Project number is required" });

        var settings = _settings.Load();
        if (string.IsNullOrWhiteSpace(settings.ProjectsRoot))
            return BadRequest(new { error = "Projects root folder is not configured. Open the admin page (/admin) to set it." });

        DirectoryInfo dir;
        try { dir = ProjectPathResolver.Resolve(number, settings.ProjectsRoot); }
        catch (Exception ex) { return BadRequest(new { error = ex.Message }); }

        if (!dir.Exists)
        {
            return Ok(new
            {
                exists       = false,
                expectedPath = dir.FullName,
                projectNumber = number
            });
        }

        var (project, employees) = _data.Load(dir.FullName);

        return Ok(new
        {
            exists        = true,
            folderPath    = dir.FullName,
            projectNumber = number,
            projectName   = project?.ProjectName?.Trim() ?? "",
            projectNotes  = project?.ProjectNotes ?? "",
            employees     = employees.Select(e => new
            {
                displayName = $"{e.FirstName} {e.LastName}".Trim(),
                email       = e.EmailAddress ?? "",
                isLeader    = e.IsLeader
            }).Where(e => !string.IsNullOrEmpty(e.email))
        });
    }
}
