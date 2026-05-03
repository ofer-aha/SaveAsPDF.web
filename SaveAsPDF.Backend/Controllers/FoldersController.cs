using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("api/project/{number}/folders")]
public class FoldersController : ControllerBase
{
    private const int MaxDepth = 5;
    private readonly SettingsService _settings;

    public FoldersController(SettingsService settings) => _settings = settings;

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
        if (ContainsInvalidChars(req.Name))
            return BadRequest(new { error = "Folder name contains invalid characters" });

        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });

        var parentPath = ResolveSafe(root, req.Parent);
        if (parentPath == null) return BadRequest(new { error = "Invalid parent path" });
        if (!Directory.Exists(parentPath)) return NotFound(new { error = "Parent folder does not exist" });

        var newFolder = Path.Combine(parentPath, req.Name);
        if (Directory.Exists(newFolder))
            return Conflict(new { error = "A folder with that name already exists" });

        Directory.CreateDirectory(newFolder);
        return Ok(new { path = ToRelative(root, newFolder), name = req.Name });
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
        if (ContainsInvalidChars(req.NewName))
            return BadRequest(new { error = "New name contains invalid characters" });

        var root = ResolveRoot(number, out var err);
        if (root == null) return BadRequest(new { error = err });

        var src = ResolveSafe(root, req.Path);
        if (src == null) return BadRequest(new { error = "Invalid source path" });
        if (string.Equals(src, root, StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Cannot rename the project root" });
        if (!Directory.Exists(src)) return NotFound(new { error = "Folder does not exist" });

        var dst = Path.Combine(Path.GetDirectoryName(src)!, req.NewName);
        if (Directory.Exists(dst))
            return Conflict(new { error = "A folder with that name already exists" });

        Directory.Move(src, dst);
        return Ok(new { path = ToRelative(root, dst), name = req.NewName });
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
        if (target == null) return BadRequest(new { error = "Invalid path" });
        if (string.Equals(target, root, StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { error = "Cannot delete the project root" });
        if (!Directory.Exists(target)) return NotFound(new { error = "Folder does not exist" });

        try
        {
            Directory.Delete(target, recursive: req.Recursive);
        }
        catch (IOException ex)
        {
            // Likely "Directory not empty" when recursive=false
            return Conflict(new { error = ex.Message });
        }
        return Ok(new { deleted = req.Path });
    }

    // ---------------------------------------------------------------
    // helpers
    // ---------------------------------------------------------------
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

    private static bool ContainsInvalidChars(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            if (name.Contains(c)) return true;
        return name == "." || name == ".." || name.StartsWith("..");
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

    public class CreateFolderRequest { public string? Parent { get; set; } public string? Name { get; set; } }
    public class RenameFolderRequest { public string? Path   { get; set; } public string? NewName { get; set; } }
    public class DeleteFolderRequest { public string? Path   { get; set; } public bool Recursive { get; set; } }
}
