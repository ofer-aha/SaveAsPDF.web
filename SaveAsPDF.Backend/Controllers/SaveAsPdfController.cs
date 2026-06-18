using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/saveaspdf")]
public class SaveAsPdfController : ControllerBase
{
    private readonly SettingsService    _settings;
    private readonly ProjectDataService _data;
    private readonly LogService         _log;

    public SaveAsPdfController(SettingsService settings, ProjectDataService data, LogService log)
    {
        _settings = settings;
        _data     = data;
        _log      = log;
    }

    private static string? ResolveSafeSubfolder(string projectRoot, string relativePath)
    {
        // Reject any input the client shouldn't send: rooted paths, alternate
        // data streams (Windows ":" sequences), or NUL bytes.
        if (string.IsNullOrWhiteSpace(relativePath)) return null;
        if (relativePath.Contains('\0') || relativePath.Contains(':')) return null;

        relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
        if (Path.IsPathRooted(relativePath)) return null;

        var combined = Path.GetFullPath(Path.Combine(projectRoot, relativePath));
        var rootFull = Path.GetFullPath(projectRoot);

        // Use GetRelativePath as the authoritative containment check — handles
        // case folding, trailing-separator quirks, and "..\" traversal cleanly.
        // A path inside the root yields a relative path that is "." or doesn't
        // start with ".." and isn't absolute.
        var rel = Path.GetRelativePath(rootFull, combined);
        if (rel == "." ) return combined;
        if (rel.StartsWith("..", StringComparison.Ordinal)) return null;
        if (Path.IsPathRooted(rel)) return null;

        return combined;
    }

    [HttpPost]
    [RequestSizeLimit(100_000_000)]
    public IActionResult Save([FromBody] SaveAsPdfRequest request)
    {
        if (request == null || string.IsNullOrWhiteSpace(request.ProjectId))
            return BadRequest("ProjectId is required");

        var settings = _settings.Load();
        var root = settings.ProjectsRoot;
        if (string.IsNullOrWhiteSpace(root))
            return BadRequest("Projects root is not configured. Open /admin to set it.");

        // Enforce admin policy server-side. If admin forced "no stamp at all",
        // null out the request's stamp before PDF generation.
        var policy = settings.StampPolicy ?? new StampPolicy();
        if (policy.DefaultStamp == false)
        {
            request.Stamp = null;
        }
        else if (request.Stamp != null)
        {
            request.Stamp.PolicyApplied = policy.ApplyTo(request.Stamp);
        }
        else if (policy.DefaultStamp == true)
        {
            // Admin forces stamp ON but client didn't send one — synthesize a default.
            request.Stamp = new StampInfo { PolicyApplied = true };
            policy.ApplyTo(request.Stamp);
        }

        // 1. Resolve project directory
        var projectDir = ProjectPathResolver.Resolve(request.ProjectId, root);
        Directory.CreateDirectory(projectDir.FullName);

        // Ensure .SaveAsPDF hidden folder exists before anything else so it is
        // created even if PDF generation or attachment saving later throws.
        try { _data.EnsureInitialized(projectDir.FullName, request.ProjectId); } catch { }

        // 2. Apply optional destination subfolder, validated against project root
        var saveDir = projectDir.FullName;
        if (!string.IsNullOrWhiteSpace(request.DestinationFolder))
        {
            var resolved = ResolveSafeSubfolder(projectDir.FullName, request.DestinationFolder);
            if (resolved == null)
                return BadRequest(new { error = "Invalid destination folder" });
            Directory.CreateDirectory(resolved);
            saveDir = resolved;
        }

        // 3. Generate PDF (mandatory) — stamp + forward info embedded in the PDF.
        //    Effective PDF settings = admin defaults + user choices, with the
        //    admin PdfPolicy forcing any locked field.
        var effectivePdf = (settings.PdfPolicy ?? new PdfPolicy())
                               .Resolve(settings.PdfSettings ?? new PdfSettings(), request.PdfSettings);
        var pdfInfo = PdfService.GeneratePdf(request.Email, saveDir, request, effectivePdf);

        // 4. Save attachments — best-effort; a corrupt attachment must never block
        //    PDF delivery or metadata persistence (step 5).
        try { AttachmentService.SaveAttachments(request.Attachments, saveDir); }
        catch (Exception ex) { Console.Error.WriteLine($"[SaveAsPDF] attachment save failed: {ex.Message}"); }

        // 5. Persist project metadata to the hidden .SaveAsPDF folder
        try
        {
            var projectModel = new ProjectXmlModel
            {
                ProjectNumber = request.ProjectId,
                ProjectName   = request.ProjectName,
                ProjectDate   = DateTime.Now.ToString("dd/MM/yyyy HH:mm"),
                LastSavePath  = pdfInfo.FullPath
            };
            var employeeModels = (request.Employees ?? []).Select((e, i) =>
            {
                var parts     = (e.DisplayName ?? "").Split(' ', 2);
                return new EmployeeXmlModel
                {
                    Id           = i + 1,
                    FirstName    = parts[0],
                    LastName     = parts.Length > 1 ? parts[1] : "",
                    EmailAddress = e.Email,
                    IsLeader     = e.IsLeader
                };
            }).ToList();
            _data.Save(projectDir.FullName, projectModel, employeeModels);
        }
        catch (Exception ex) { Console.Error.WriteLine($"[SaveAsPDF] metadata save failed: {ex.Message}"); }

        // 6. Log the save event (best-effort — must never block the response)
        try
        {
            _log.Log(new LogEntry
            {
                Level       = "success",
                Username    = request.SavedBy ?? request.Stamp?.UserName ?? "",
                Subject     = request.Email?.Subject ?? "",
                Attachments = (request.Attachments?.Select(a => a.Name ?? "")
                               ?? Enumerable.Empty<string>()).ToArray(),
                ProjectId   = request.ProjectId!,
                ProjectName = request.ProjectName ?? "",
                SavePath    = pdfInfo.FullPath
            });
        }
        catch (Exception ex) { Console.Error.WriteLine($"[SaveAsPDF] log failed: {ex.Message}"); }

        // 7. Return result
        return Ok(new
        {
            pdf = pdfInfo,
            project = new
            {
                request.ProjectId,
                FullName = projectDir.FullName,
                SaveDir  = saveDir
            }
        });
    }


}