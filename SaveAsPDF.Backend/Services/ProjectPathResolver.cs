using System;
using System.IO;
using System.Text.RegularExpressions;

public static class ProjectPathResolver
{
    public static DirectoryInfo Resolve(string projectNumber, string rootDrive)
    {
        if (string.IsNullOrWhiteSpace(projectNumber))
            throw new ArgumentException("Project number cannot be empty or invalid.", nameof(projectNumber));

        if (string.IsNullOrWhiteSpace(rootDrive))
            throw new ArgumentException("Root drive cannot be empty or invalid.", nameof(rootDrive));

        projectNumber = NormalizeProjectNumber(projectNumber);
        string normalizedRoot = NormalizeRootPath(rootDrive);

        if (!projectNumber.SafeProjectID())
            return new DirectoryInfo(normalizedRoot);

        string[] split = projectNumber.Split('-');
        string firstPart = split[0];
        string level1 = GetLevel1Folder(firstPart);

        string level1Path = Path.Combine(normalizedRoot, level1);

        // Legacy flat layout under root:
        // <root>\1000 or <root>\1000-1
        string rootFlatProjectPath = Path.Combine(normalizedRoot, projectNumber);
        if (Directory.Exists(rootFlatProjectPath))
            return new DirectoryInfo(rootFlatProjectPath);

        // <root>\10\1000
        if (split.Length == 1)
        {
            return new DirectoryInfo(Path.Combine(level1Path, firstPart));
        }

        // <root>\10\1000-1
        string flatProjectPath = Path.Combine(level1Path, projectNumber);

        // <root>\10\1000\1000-1
        string nestedBaseByFirst = Path.Combine(level1Path, firstPart);
        string nestedProjectByFirst = Path.Combine(nestedBaseByFirst, projectNumber);

        // Legacy nested under root base:
        // <root>\1000\1000-1
        string nestedProjectUnderRootBase =
            Path.Combine(Path.Combine(normalizedRoot, firstPart), projectNumber);

        // For 2+ hyphens:
        // 1000-2-1 => <root>\10\1000-2\1000-2-1
        string nestedBaseByFirstTwo = split.Length > 2
            ? Path.Combine(level1Path, firstPart + "-" + split[1])
            : string.Empty;

        string nestedProjectByFirstTwo = !string.IsNullOrEmpty(nestedBaseByFirstTwo)
            ? Path.Combine(nestedBaseByFirstTwo, projectNumber)
            : string.Empty;

        // Prefer exact existing project folder (most specific first)
        if (!string.IsNullOrEmpty(nestedProjectByFirstTwo) &&
            Directory.Exists(nestedProjectByFirstTwo))
            return new DirectoryInfo(nestedProjectByFirstTwo);

        if (Directory.Exists(nestedProjectByFirst))
            return new DirectoryInfo(nestedProjectByFirst);

        if (Directory.Exists(nestedProjectUnderRootBase))
            return new DirectoryInfo(nestedProjectUnderRootBase);

        if (Directory.Exists(flatProjectPath))
            return new DirectoryInfo(flatProjectPath);

        // Fall back to existing bases
        if (!string.IsNullOrEmpty(nestedBaseByFirstTwo) &&
            Directory.Exists(nestedBaseByFirstTwo))
            return new DirectoryInfo(nestedProjectByFirstTwo);

        if (Directory.Exists(nestedBaseByFirst))
            return new DirectoryInfo(nestedProjectByFirst);

        if (Directory.Exists(Path.Combine(normalizedRoot, firstPart)))
            return new DirectoryInfo(nestedProjectUnderRootBase);

        // Default: new project under first-part base
        return new DirectoryInfo(nestedProjectByFirst);
    }

    // ------------------------
    // helpers
    // ------------------------

    private static string NormalizeProjectNumber(string projectNumber)
    {
        if (string.IsNullOrWhiteSpace(projectNumber))
            return string.Empty;

        var normalized = projectNumber.Trim();
        normalized = Regex.Replace(normalized, @"\s*-\s*", "-");
        normalized = Regex.Replace(normalized, @"\s+", string.Empty);
        return normalized;
    }

    private static string NormalizeRootPath(string rootDrive)
    {
        var root = rootDrive.Trim().Replace('/', '\\');
        root = EnsureDriveRoot(root);
        root = TryNormalizePath(root);
        return EnsureDriveRoot(root);
    }

    private static string EnsureDriveRoot(string path)
    {
        if (!path.EndsWith("\\"))
            path += "\\";
        return path;
    }

    private static string TryNormalizePath(string path)
    {
        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
    }

    // Level-1 folder = floor(numericBase / 100), zero-padded to 2 digits.
    //   123     → 1   → "01"
    //   1000    → 10  → "10"
    //   1234    → 12  → "12"
    //   1234A   → 12  → "12"  (alpha suffix ignored for level1)
    private static string GetLevel1Folder(string firstPart)
    {
        var digits = Regex.Match(firstPart, @"^\d+");
        if (!digits.Success) return "00";
        if (!int.TryParse(digits.Value, out var n)) return "00";
        return (n / 100).ToString("D2");
    }

}