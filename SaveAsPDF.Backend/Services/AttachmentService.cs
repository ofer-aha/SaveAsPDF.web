public static class AttachmentService
{
    public static void SaveAttachments(
        List<AttachmentDto>? attachments,
        string saveDir)
    {
        if (attachments == null || attachments.Count == 0)
            return;

        Directory.CreateDirectory(saveDir);

        foreach (var att in attachments)
        {
            if (string.IsNullOrEmpty(att?.Name) || string.IsNullOrEmpty(att.Base64))
                continue;

            var filePath = ResolveUniquePath(saveDir, att.Name);
            var bytes    = Convert.FromBase64String(att.Base64);
            File.WriteAllBytes(filePath, bytes);
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
