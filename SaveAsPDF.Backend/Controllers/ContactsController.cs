using Microsoft.AspNetCore.Mvc;
using System.Data.OleDb;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Runtime.Versioning;
using System.Text.RegularExpressions;

namespace SaveAsPDF.Backend.Controllers;

[ApiController]
[Route("api/contacts")]
[SupportedOSPlatform("windows")]
public class ContactsController : ControllerBase
{
    // Windows Search indexes all Outlook contacts automatically — no COM, no auth needed.
    private const string WdsConnection =
        "Provider=Search.CollatorDSO;Extended Properties='Application=Windows Search'";

    // Windows Search accepts either property name form; we try both variants.
    private static readonly string[] EmailColumnCandidates =
    [
        "System.Contact.PrimaryEmail",
        "System.Contact.EmailAddress",
        "System.Contact.PrimaryEmailAddress",
    ];

    [HttpGet]
    public async Task<IActionResult> GetContacts()
    {
        try
        {
            var data = await FetchFromWindowsSearch();
            return Ok(data);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    private static async Task<ContactData> FetchFromWindowsSearch()
    {
        using var conn = new OleDbConnection(WdsConnection);
        await conn.OpenAsync();

        // Discover which email column name Windows Search accepts on this machine
        string emailCol = await ResolveEmailColumn(conn);

        string sql = $@"
            SELECT
                ""System.Contact.FullName"",
                ""{emailCol}"",
                ""System.ItemFolderPathDisplay""
            FROM SYSTEMINDEX
            WHERE ""System.Kind"" = 'contact'
            ORDER BY ""System.Contact.FullName""";

        var byPath = new Dictionary<string, List<ContactDto>>(StringComparer.OrdinalIgnoreCase);

        using var cmd    = new OleDbCommand(sql, conn);
        using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            string name  = Str(reader, "System.Contact.FullName");
            string email = Str(reader, emailCol);
            string path  = Str(reader, "System.ItemFolderPathDisplay");

            if (string.IsNullOrWhiteSpace(name) || !email.Contains('@')) continue;

            if (!byPath.ContainsKey(path))
                byPath[path] = [];

            byPath[path].Add(new ContactDto(name, email));
        }

        // Merge domain users from LDAP (may be empty if not domain-joined)
        var domainUsers = FetchDomainUsers();
        if (domainUsers.Count > 0)
            byPath["__domain__/כל משתמשי הדומיין"] = domainUsers;

        return BuildTree(byPath);
    }

    private static List<ContactDto> FetchDomainUsers()
    {
        var result = new List<ContactDto>();
        try
        {
            // GetCurrentDomain() correctly resolves dw.local / any Windows domain
            var domainEntry = Domain.GetCurrentDomain().GetDirectoryEntry();

            using var searcher = new DirectorySearcher(domainEntry)
            {
                Filter   = "(&(objectClass=user)(mail=*)(!(objectClass=computer))" +
                           "(!(userAccountControl:1.2.840.113556.1.4.803:=2)))",
                PageSize  = 1000,
                SizeLimit = 5000
            };
            searcher.PropertiesToLoad.AddRange(new[] { "displayName", "mail" });

            foreach (SearchResult r in searcher.FindAll())
            {
                string name  = r.Properties["displayName"]?.Count > 0
                    ? r.Properties["displayName"][0]?.ToString()?.Trim() ?? "" : "";
                string email = r.Properties["mail"]?.Count > 0
                    ? r.Properties["mail"][0]?.ToString()?.Trim() ?? "" : "";

                if (!string.IsNullOrEmpty(name) && email.Contains('@'))
                    result.Add(new ContactDto(name, email));
            }

            result.Sort((a, b) =>
                string.Compare(a.DisplayName, b.DisplayName, StringComparison.CurrentCulture));
        }
        catch { /* LDAP unavailable — silently skip */ }

        return result;
    }

    // Try each candidate column name with a cheap query until one succeeds.
    private static async Task<string> ResolveEmailColumn(OleDbConnection conn)
    {
        foreach (var col in EmailColumnCandidates)
        {
            try
            {
                var probe = new OleDbCommand(
                    $@"SELECT TOP 1 ""{col}"" FROM SYSTEMINDEX WHERE ""System.Kind"" = 'contact'",
                    conn);
                await probe.ExecuteReaderAsync();
                return col;
            }
            catch { }
        }
        throw new Exception(
            $"No known email column found in Windows Search. Tried: {string.Join(", ", EmailColumnCandidates)}");
    }

    // Build a parent/child folder tree from raw folder paths.
    // Paths look like: "\\mailbox@domain\Contacts\Work" or a file-system path.
    private static ContactData BuildTree(Dictionary<string, List<ContactDto>> byPath)
    {
        var folders  = new List<FolderDto>();
        var contacts = new Dictionary<string, List<ContactDto>>();

        // Sort shortest-first so parents are registered before children
        var sortedPaths = byPath.Keys.OrderBy(p => p.Length).ThenBy(p => p).ToList();
        var pathToId    = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < sortedPaths.Count; i++)
        {
            string path = sortedPaths[i];
            string id   = $"f{i}";
            pathToId[path] = id;

            // Nearest registered ancestor = parent
            string? parentId = null;
            string? bestAncestor = null;
            foreach (var (p, pid) in pathToId)
            {
                if (p == path) continue;
                if (path.StartsWith(p, StringComparison.OrdinalIgnoreCase) &&
                    (bestAncestor == null || p.Length > bestAncestor.Length))
                {
                    bestAncestor = p;
                    parentId = pid;
                }
            }

            // Display name = last non-empty segment
            string displayName = path
                .Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault() ?? path;

            // Skip hidden / system folders
            if (displayName.StartsWith("hiddenFolder", StringComparison.OrdinalIgnoreCase) ||
                Regex.IsMatch(displayName, @"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}",
                              RegexOptions.IgnoreCase))
                continue;

            folders.Add(new FolderDto(id, displayName, parentId));
            contacts[id] = byPath[path];
        }

        return new ContactData(folders, contacts);
    }

    private static string Str(System.Data.Common.DbDataReader r, string col)
    {
        try { return r[col]?.ToString()?.Trim() ?? ""; }
        catch { return ""; }
    }

    record FolderDto(string Id, string DisplayName, string? ParentId);
    record ContactDto(string DisplayName, string Email);
    record ContactData(List<FolderDto> Folders, Dictionary<string, List<ContactDto>> Contacts);
}
