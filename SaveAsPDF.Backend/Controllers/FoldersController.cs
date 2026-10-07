using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/project/{number}/folders")]
public class FoldersController : ControllerBase
{
    private const int MaxDepth = 5;
    private readonly SettingsService _settings;
    private readonly LogService      _log;

    public FoldersController(SettingsService settings, LogService log)
    {
        _settings = settings;
        _log      = log;
    }

    // ---------------------------------------------------------------
    // GET /api/project/{number}/folders/tree
    // ---------------------------------------------------------------
    [HttpGet("tree")]
    public IActionResult GetTree(string number)
    {
        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });
        if (!Directory.Exists(root))
            return NotFound(new { error = "Project folder does not exist", expectedPath = root });

        return Ok(new { rootPath = root, tree = BuildNode(root, "", 0) });
    }

    // ---------------------------------------------------------------
    // POST /api/project/{number}/folders   { parent, name }
    // ---------------------------------------------------------------
    [HttpPost]
    public IActionResult Create(string number, [FromBody] CreateFolderRequest req)
    {
        if (req == null || string.IsNullOrWhiteSpace(req.Name))
            return BadRequest(new { error = "Name is required" });

        // Cosmetic problems are fixed silently; escape attempts are refused and logged.
        var check = FolderNameRules.Check(req.Name);
        if (!check.Ok)
        {
            if (check.IsSecurity) LogSecurityEvent(number, "create", req.Name, check, req.User);
            return BadRequest(new { error = check.Message, code = check.Code });
        }
        var safeName = check.Value;

        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });

        var parentPath = ResolveSafe(root, req.Parent);
        if (parentPath == null)
        {
            LogSecurityEvent(number, "create", req.Parent ?? "", null, req.User, "Parent path escapes the project root");
            return BadRequest(new { error = "Invalid parent path" });
        }
        if (!Directory.Exists(parentPath)) return NotFound(new { error = "Parent folder does not exist" });

        var newFolder = Path.Combine(parentPath, safeName);

        // Belt and braces: the name passed the rules, so this cannot normally fail -
        // but a path that resolves outside the root must never be created.
        if (ResolveSafe(root, ToRelative(root, newFolder)) == null)
        {
            LogSecurityEvent(number, "create", req.Name, null, req.User, "Resolved path escapes the project root");
            return BadRequest(new { error = "Invalid folder name" });
        }

        if (Directory.Exists(newFolder))
            return Conflict(new { error = "A folder with that name already exists" });

        try { Directory.CreateDirectory(newFolder); }
        catch (Exception ex) { return FileSystemError(ex); }
        return Ok(new { path = ToRelative(root, newFolder), name = safeName });
    }

    // ---------------------------------------------------------------
    // PUT /api/project/{number}/folders   { path, newName }
    // ---------------------------------------------------------------
    [HttpPut]
    public IActionResult Rename(string number, [FromBody] RenameFolderRequest req)
    {
        if (req == null || string.IsNullOrWhiteSpace(req.Path))
            return BadRequest(new { error = "Path is required" });
        if (string.IsNullOrWhiteSpace(req.NewName))
            return BadRequest(new { error = "New name is required" });

        var check = FolderNameRules.Check(req.NewName);
        if (!check.Ok)
        {
            if (check.IsSecurity) LogSecurityEvent(number, "rename", req.NewName, check, req.User);
            return BadRequest(new { error = check.Message, code = check.Code });
        }
        var safeName = check.Value;

        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });

        var src = ResolveSafe(root, req.Path);
        if (src == null)
        {
            LogSecurityEvent(number, "rename", req.Path ?? "", null, req.User, "Source path escapes the project root");
            return BadRequest(new { error = "Invalid source path" });
        }
        if (string.Equals(src, root, StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Cannot rename the project root" });
        if (!Directory.Exists(src)) return NotFound(new { error = "Folder does not exist" });

        var dst = Path.Combine(Path.GetDirectoryName(src)!, safeName);
        if (ResolveSafe(root, ToRelative(root, dst)) == null)
        {
            LogSecurityEvent(number, "rename", req.NewName, null, req.User, "Resolved path escapes the project root");
            return BadRequest(new { error = "Invalid folder name" });
        }
        if (Directory.Exists(dst))
            return Conflict(new { error = "A folder with that name already exists" });

        try { Directory.Move(src, dst); }
        catch (Exception ex) { return FileSystemError(ex); }
        return Ok(new { path = ToRelative(root, dst), name = safeName });
    }

    // ---------------------------------------------------------------
    // DELETE /api/project/{number}/folders   { path, recursive }
    // ---------------------------------------------------------------
    [HttpDelete]
    public IActionResult Delete(string number, [FromBody] DeleteFolderRequest req)
    {
        if (req == null || string.IsNullOrWhiteSpace(req.Path))
            return BadRequest(new { error = "Path is required" });

        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });

        var target = ResolveSafe(root, req.Path);
        if (target == null)
        {
            LogSecurityEvent(number, "delete", req.Path ?? "", null, req.User, "Path escapes the project root");
            return BadRequest(new { error = "Invalid path" });
        }
        if (string.Equals(target, root, StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Cannot delete the project root" });
        if (!Directory.Exists(target)) return NotFound(new { error = "Folder does not exist" });

        try { Directory.Delete(target, recursive: req.Recursive); }
        catch (Exception ex) { return FileSystemError(ex); }
        return Ok(new { deleted = req.Path });
    }

    // ---------------------------------------------------------------
    // POST /api/project/{number}/folders/report-blocked   { name, action, user }
    // ---------------------------------------------------------------
    // The taskpane blocks these names before they are ever submitted, which is
    // good UX but would leave no trace. It reports them here so the attempt is
    // still logged. The value is re-validated server-side, so this endpoint
    // cannot be used to write arbitrary text into the log.
    [HttpPost("report-blocked")]
    public IActionResult ReportBlocked(string number, [FromBody] ReportBlockedRequest req)
    {
        if (req == null || string.IsNullOrWhiteSpace(req.Name))
            return BadRequest(new { error = "Name is required" });

        var check = FolderNameRules.Check(req.Name);
        if (check.Ok || !check.IsSecurity)
            return Ok(new { logged = false });   // not actually a violation - ignore

        var action = req.Action == "rename" ? "rename" : "create";
        LogSecurityEvent(number, action, req.Name, check, req.User);
        return Ok(new { logged = true });
    }

    // ---------------------------------------------------------------
    // helpers
    // ---------------------------------------------------------------

    /// <summary>
    /// Turns a filesystem exception into a status code and a Hebrew message the
    /// user can act on.
    /// </summary>
    /// <remarks>
    /// These calls used to be unguarded, and there is no exception-handler
    /// middleware, so the most ordinary situation on a shared drive - the folder
    /// is open in Explorer or Word - produced a 500 with an empty body and the
    /// taskpane showed the user the bare number "500". The fix the user needs is
    /// "close the window you have open", which the message now says.
    /// </remarks>
    private ObjectResult FileSystemError(Exception ex) => ex switch
    {
        UnauthorizedAccessException => StatusCode(403, new
        {
            error = "אין הרשאה לבצע את הפעולה על התיקיה הזו.",
            code  = "denied"
        }),
        DirectoryNotFoundException => StatusCode(404, new
        {
            error = "התיקיה לא נמצאה - ייתכן שמישהו אחר מחק או העביר אותה.",
            code  = "not_found"
        }),
        PathTooLongException => StatusCode(400, new
        {
            error = "הנתיב ארוך מדי. קצר/י את שם התיקיה או שמור/י ברמה גבוהה יותר.",
            code  = "path_too_long"
        }),
        IOException => StatusCode(409, new
        {
            error = "התיקיה בשימוש או אינה ריקה. סגור/י חלונות פתוחים ונסה/י שוב.",
            code  = "in_use"
        }),
        _ => StatusCode(500, new
        {
            error = "הפעולה נכשלה: " + ex.Message,
            code  = "error"
        })
    };

    private string? ResolveRoot(string number, out string? error)
    {
        error = null;
        var settings = _settings.Load();
        if (string.IsNullOrWhiteSpace(settings.ProjectsRoot))
        {
            error = "Projects root is not configured. Open /admin to set it.";
            return null;
        }
        try
        {
            var dir = ProjectPathResolver.Resolve(number, settings.ProjectsRoot);
            return dir.FullName;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return null;
        }
    }

    // Resolves a relative subpath against the project root and ensures the
    // result is contained within it (defends against ".." traversal,
    // alternate data streams, NUL injection).
    private static string? ResolveSafe(string projectRoot, string? relativePath)
    {
        relativePath ??= "";
        if (relativePath.Contains('\0') || relativePath.Contains(':')) return null;

        relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
        if (Path.IsPathRooted(relativePath)) return null;

        var combined = Path.GetFullPath(Path.Combine(projectRoot, relativePath));
        var rootFull = Path.GetFullPath(projectRoot);

        var rel = Path.GetRelativePath(rootFull, combined);
        if (rel == ".") return combined;
        if (rel.StartsWith("..", StringComparison.Ordinal)) return null;
        if (Path.IsPathRooted(rel)) return null;

        return combined;
    }

    private static string ToRelative(string root, string fullPath)
    {
        var rel = Path.GetRelativePath(root, fullPath);
        return rel.Replace(Path.DirectorySeparatorChar, '/');
    }

    // Every refused escape attempt lands in the same Save Events log the admin
    // already watches, so a probe against this API is visible rather than silent.
    // The attempted value is truncated - a log entry must never become the payload.
    private void LogSecurityEvent(
        string  projectNumber,
        string  action,
        string  attempted,
        FolderNameRules.Result? check,
        string? user,
        string? overrideMessage = null)
    {
        try
        {
            var reason = overrideMessage ?? check?.Message ?? "Blocked folder path";
            var code   = check?.Code ?? "traversal";
            var shown  = attempted.Length > 200 ? attempted[..200] + "..." : attempted;
            var ip     = HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? "unknown";

            _log.Log(new LogEntry
            {
                Level       = "error",
                Username    = string.IsNullOrWhiteSpace(user) ? "(unknown)" : user!,
                Subject     = $"Blocked folder {action}: {shown}",
                ProjectId   = projectNumber ?? "",
                ProjectName = "",
                SavePath    = null,
                ErrorDetail = $"SECURITY [{code}] {reason}. Attempted value: {shown}. Client: {ip}."
            });
        }
        catch { /* logging must never break the request path */ }
    }

    private static FolderNode BuildNode(string fullPath, string relativePath, int depth)
    {
        var node = new FolderNode
        {
            Name = depth == 0 ? Path.GetFileName(fullPath.TrimEnd(Path.DirectorySeparatorChar)) : Path.GetFileName(fullPath),
            Path = relativePath,
            Children = new List<FolderNode>()
        };

        if (depth >= MaxDepth) return node;

        try
        {
            foreach (var sub in Directory.EnumerateDirectories(fullPath))
            {
                var name = Path.GetFileName(sub);
                if (string.IsNullOrEmpty(name)) continue;
                if (name.StartsWith(".")) continue; // hide dotfolders (e.g. .SaveAsPDF)

                var childRel = string.IsNullOrEmpty(relativePath)
                    ? name
                    : relativePath + "/" + name;
                node.Children.Add(BuildNode(sub, childRel, depth + 1));
            }
        }
        catch { /* permission errors on a subtree shouldn't kill the whole response */ }

        node.Children.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
        return node;
    }

    public class FolderNode
    {
        public string Name     { get; set; } = "";
        public string Path     { get; set; } = "";
        public List<FolderNode> Children { get; set; } = new();
    }

    // User is best-effort context for the audit log (the taskpane sends the
    // Outlook display name); it is never used for authorization.
    public class CreateFolderRequest { public string? Parent { get; set; } public string? Name { get; set; } public string? User { get; set; } }
    public class RenameFolderRequest { public string? Path   { get; set; } public string? NewName { get; set; } public string? User { get; set; } }
    public class DeleteFolderRequest { public string? Path   { get; set; } public bool Recursive { get; set; } public string? User { get; set; } }
    public class ReportBlockedRequest { public string? Name  { get; set; } public string? Action  { get; set; } public string? User { get; set; } }
}
