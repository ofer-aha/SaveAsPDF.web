using System.Text.RegularExpressions;

public static class ProjectIdExtensions
{
    public static bool SafeProjectID(this string projectId)
    {
        // 3-4 digit base, optional single trailing alpha char (XXXY),
        // followed by zero-or-more "-<alphanumeric>" segments.
        return Regex.IsMatch(projectId, @"^\d{3,4}[A-Za-z]?(-[A-Za-z0-9]+)*$");
    }
}
