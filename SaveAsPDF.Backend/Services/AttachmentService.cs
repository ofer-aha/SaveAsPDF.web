public static class AttachmentService
{
    public static void SaveAttachments(
        List<AttachmentDto>? attachments,
        string saveDir)
    {
        if (attachments == null || attachments.Count == 0)
            return;

        try { Directory.CreateDirectory(saveDir); } catch { return; }

        foreach (var att in attachments)
        {
            if (string.IsNullOrEmpty(att?.Name) || string.IsNullOrEmpty(att.Base64))
                continue;

            try
            {
                var clean = att.Base64.Replace("\r", "").Replace("\n", "").Replace("\t", "").Replace(" ", "");
                var bytes = Convert.FromBase64String(clean);
                var filePath = ResolveUniquePath(saveDir, att.Name);
                File.WriteAllBytes(filePath, bytes);
            }
            catch { /* skip individual corrupt attachments — PDF save must not fail */ }
        }
    }

    // Returns a path that does not yet exist by appending (1), (2)... before the extension
    // when the original name is taken: file.txt -> file(1).txt -> file(2).txt ...
    private static string ResolveUniquePath(string dir, string originalName)
    {
        var safe = SanitizeFileName(originalName);
        var candidate = Path.Combine(dir, safe);
        if (!File.Exists(candidate)) return candidate;

        var stem = Path.GetFileNameWithoutExtension(safe);
        var ext  = Path.GetExtension(safe); // includes the leading dot, or "" if none

        for (var i = 1; i < 10000; i++)
        {
            candidate = Path.Combine(dir, $"{stem}({i}){ext}");
            if (!File.Exists(candidate)) return candidate;
        }
        // Extreme fallback — append a guid to guarantee uniqueness
        return Path.Combine(dir, $"{stem}_{Guid.NewGuid():N}{ext}");
    }

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}
