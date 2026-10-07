public static class AttachmentService
{
    // Ceilings, not expectations. The endpoint writes caller-supplied bytes into a
    // shared project folder, so without a cap a single request can fill the file
    // server. Real e-mails are far below these numbers; anything above is either a
    // mistake or abuse.
    private const int  MaxAttachmentCount = 50;
    private const long MaxTotalBytes      = 80L * 1024 * 1024;   // 80 MB per save

    public static void SaveAttachments(
        List<AttachmentDto>? attachments,
        string saveDir)
    {
        if (attachments == null || attachments.Count == 0)
            return;

        try { Directory.CreateDirectory(saveDir); } catch { return; }

        var written = 0;
        long totalBytes = 0;

        foreach (var att in attachments)
        {
            if (string.IsNullOrEmpty(att?.Name) || string.IsNullOrEmpty(att.Base64))
                continue;

            if (written >= MaxAttachmentCount)
            {
                Console.Error.WriteLine(
                    $"[SaveAsPDF] attachment limit reached ({MaxAttachmentCount}); remaining attachments skipped.");
                break;
            }

            try
            {
                var clean = att.Base64.Replace("\r", "").Replace("\n", "").Replace("\t", "").Replace(" ", "");
                var bytes = Convert.FromBase64String(clean);

                if (totalBytes + bytes.LongLength > MaxTotalBytes)
                {
                    Console.Error.WriteLine(
                        $"[SaveAsPDF] attachment size budget exhausted; skipped '{att.Name}'.");
                    continue;
                }

                var filePath = ResolveUniquePath(saveDir, att.Name);
                File.WriteAllBytes(filePath, bytes);
                totalBytes += bytes.LongLength;
                written++;
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

    // Attachment names come straight from an email, so they are attacker-chosen.
    // Path.GetInvalidFileNameChars() covers separators and control chars, but NOT:
    //   - "." / ".."            -> resolve to the directory itself or its parent
    //   - trailing dots/spaces  -> Windows silently strips them, so "evil.txt " and
    //                              "evil.txt" collide and can overwrite each other
    //   - reserved device names -> CON, PRN, AUX, NUL, COM1-9, LPT1-9 (with or
    //                              without an extension) still resolve to devices
    //   - very long names       -> blow the 260-char path limit and throw
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "CON","PRN","AUX","NUL",
        "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
        "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
    };

    private static string SanitizeFileName(string name)
    {
        foreach (var c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');

        name = name.TrimEnd('.', ' ');

        if (string.IsNullOrWhiteSpace(name) || name == "." || name == "..")
            return "attachment";

        var stem = Path.GetFileNameWithoutExtension(name);
        if (ReservedNames.Contains(stem))
            name = "_" + name;

        // Leave room for the "(nn)" uniqueness suffix and the parent directory.
        const int MaxLen = 120;
        if (name.Length > MaxLen)
        {
            var ext = Path.GetExtension(name);
            if (ext.Length > 20) ext = "";
            var keep = Math.Max(1, MaxLen - ext.Length);
            name = Path.GetFileNameWithoutExtension(name);
            name = name[..Math.Min(name.Length, keep)] + ext;
        }

        return name;
    }
}
