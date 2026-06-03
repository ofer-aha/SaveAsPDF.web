using System.DirectoryServices;

public static class AdGroupService
{
    public static bool IsUserInGroup(string userEmail, string groupName)
    {
        if (string.IsNullOrWhiteSpace(userEmail) || string.IsNullOrWhiteSpace(groupName))
            return false;
        try
        {
            // Resolve group DN first
            using var groupSearcher = new DirectorySearcher
            {
                Filter = $"(&(objectClass=group)(|(cn={Esc(groupName)})(sAMAccountName={Esc(groupName)})))",
                PropertiesToLoad = { "distinguishedName" }
            };
            var groupEntry = groupSearcher.FindOne();
            if (groupEntry == null) return false;
            var groupDn = groupEntry.Properties["distinguishedName"][0]?.ToString();
            if (string.IsNullOrEmpty(groupDn)) return false;

            // Transitive membership check (LDAP_MATCHING_RULE_IN_CHAIN)
            using var userSearcher = new DirectorySearcher
            {
                Filter = $"(&(objectClass=user)(mail={Esc(userEmail)})(memberOf:1.2.840.113556.1.4.1941:={groupDn}))",
                PropertiesToLoad = { "sAMAccountName" }
            };
            return userSearcher.FindOne() != null;
        }
        catch
        {
            return false; // AD unavailable or non-domain machine — deny silently
        }
    }

    private static string Esc(string s) =>
        s.Replace("\\", "\\5C").Replace("*",  "\\2A")
         .Replace("(",  "\\28").Replace(")",  "\\29")
         .Replace("\0", "\\00");
}
