using System.Text;
using System.Text.RegularExpressions;

/// <summary>
/// Single source of truth for what a user may name a folder.
///
/// Two very different failure modes are deliberately kept apart:
///
///   * <b>Cosmetic</b> - characters Windows simply cannot store in a name
///     (&lt; &gt; : " | ? * and control chars). These are silently stripped;
///     nagging the user about a stray colon helps nobody.
///
///   * <b>Blocked</b> - input that is trying to leave the project folder
///     (path separators, "..", drive letters, UNC, percent-encoded traversal)
///     or that names a Windows device (CON, PRN, NUL, COM1...). These are
///     refused outright, reported to the user, and logged as security events.
///
/// The taskpane mirrors these rules for instant feedback, but this class is
/// the boundary that actually enforces them - the client is not trusted.
/// </summary>
public static class FolderNameRules
{
    public const int MaxLength = 200;

    // CON, PRN, AUX, NUL, COM0-9, LPT0-9 - reserved by Windows even with an
    // extension ("CON.txt") and regardless of case.
    private static readonly Regex ReservedRe =
        new(@"^(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])(\..*)?$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // Percent-encoded separators / dot-segment. Narrow on purpose: a bare "%2e"
    // can occur in a legitimate name, an encoded separator cannot.
    private static readonly Regex EncodedTraversalRe =
        new(@"%2f|%5c|%2e%2e", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex DriveRe =
        new(@"^[A-Za-z]:", RegexOptions.Compiled);

    /// <summary>Result of checking a proposed folder name.</summary>
    /// <param name="Ok">True when the (sanitized) name may be used.</param>
    /// <param name="Value">The sanitized name. Meaningless when Ok is false.</param>
    /// <param name="Code">empty | traversal | reserved - machine-readable reason.</param>
    /// <param name="Message">English description, used for the API error and the log.</param>
    /// <param name="IsSecurity">True when the rejection is an escape attempt, not a typo.</param>
    public record Result(bool Ok, string Value, string? Code, string? Message, bool IsSecurity);

    /// <summary>Strip characters Windows cannot store, without changing intent.</summary>
    public static string Sanitize(string? raw)
    {
        if (string.IsNullOrEmpty(raw)) return "";

        var sb = new StringBuilder(raw.Length);
        foreach (var c in raw)
        {
            if (char.IsControl(c)) continue;                 // includes NUL
            if ("<>:|?*\"".IndexOf(c) >= 0) continue;        // illegal in a Windows name
            sb.Append(c);
        }

        var s = sb.ToString();
        s = Regex.Replace(s, @"\s+", " ").Trim();

        // Windows silently drops trailing dots and spaces; do it explicitly so the
        // name the user sees is the name that ends up on disk.
        s = s.TrimEnd('.', ' ');

        if (s.Length > MaxLength) s = s[..MaxLength].TrimEnd('.', ' ');
        return s;
    }

    /// <summary>Sanitize, then decide whether the result may be used.</summary>
    /// <remarks>
    /// Order matters. Every check runs on the RAW input and must run before
    /// <see cref="Sanitize"/>, which erases the evidence being looked for: it
    /// strips trailing dots, so ".." would arrive as an empty string and be
    /// reported as "no name" instead of as a traversal attempt.
    ///
    /// Note what is NOT blocked: ".." *inside* a name. "as..as" is a perfectly
    /// ordinary folder name. Dots only navigate when they form a whole path
    /// segment, and separators are refused outright, so the only dangerous case
    /// left is a name that is nothing but dots.
    /// </remarks>
    public static Result Check(string? raw)
    {
        raw ??= "";
        var trimmed = raw.Trim();

        if (raw.Contains('\0'))
            return new Result(false, "", "traversal", "Folder name may not contain a NUL character", true);

        // A name made only of dots IS a path segment: ".." resolves to the parent.
        if (trimmed.Length > 0 && trimmed.All(c => c == '.'))
            return new Result(false, "", "traversal", "Folder name may not consist only of dots", true);

        // Before the separator check: "C:\Windows" is a drive escape, not a typo.
        if (DriveRe.IsMatch(trimmed))
            return new Result(false, "", "traversal", "Folder name may not start with a drive letter", true);

        if (EncodedTraversalRe.IsMatch(raw))
            return new Result(false, "", "traversal", "Folder name contains an encoded path separator", true);

        var hasSeparator = raw.Contains('/') || raw.Contains('\\');

        // A separator together with ".." is the classic escape attempt.
        if (hasSeparator && raw.Contains(".."))
            return new Result(false, "", "traversal", "Folder name contains a path traversal sequence", true);

        // A separator on its own is a usability problem, not an attack: the user
        // is trying to create a folder and a subfolder in one step. Refused, but
        // not logged as a security event.
        if (hasSeparator)
            return new Result(false, "", "separator",
                "Folder name may not contain a path separator - create the folder first, then the subfolder inside it", false);

        var value = Sanitize(raw);

        if (value.Length == 0)
            return new Result(false, "", "empty", "Folder name is empty after removing invalid characters", false);

        if (ReservedRe.IsMatch(value))
            return new Result(false, "", "reserved",
                $"'{value}' is a name reserved by Windows and cannot be used for a folder", false);

        return new Result(true, value, null, null, false);
    }
}
