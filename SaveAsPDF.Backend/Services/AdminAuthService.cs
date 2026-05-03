using System.Security.Cryptography;
using System.Text;

public static class AdminAuthService
{
    private const int Iterations = 100_000;
    private const int SaltSize   = 16;
    private const int HashSize   = 32;

    public const string DefaultUsername = "admin";
    public const string DefaultPassword = "admin";

    public static AdminCredentials Hash(string username, string password)
    {
        var salt = RandomNumberGenerator.GetBytes(SaltSize);
        var hash = Rfc2898DeriveBytes.Pbkdf2(
            password, salt, Iterations, HashAlgorithmName.SHA256, HashSize);
        return new AdminCredentials
        {
            Username     = username,
            PasswordHash = Convert.ToBase64String(hash),
            PasswordSalt = Convert.ToBase64String(salt)
        };
    }

    public static bool Verify(AdminCredentials? creds, string username, string password)
    {
        // First-time setup: if no creds saved yet, accept the default admin/admin.
        if (creds == null
            || string.IsNullOrEmpty(creds.PasswordHash)
            || string.IsNullOrEmpty(creds.PasswordSalt))
        {
            return username == DefaultUsername && password == DefaultPassword;
        }

        if (!string.Equals(creds.Username, username, StringComparison.Ordinal))
            return false;

        try
        {
            var salt     = Convert.FromBase64String(creds.PasswordSalt);
            var expected = Convert.FromBase64String(creds.PasswordHash);
            var actual   = Rfc2898DeriveBytes.Pbkdf2(
                password, salt, Iterations, HashAlgorithmName.SHA256, HashSize);
            return CryptographicOperations.FixedTimeEquals(actual, expected);
        }
        catch
        {
            return false;
        }
    }

    public static bool IsUsingDefaults(AdminCredentials? creds) =>
        creds == null
        || string.IsNullOrEmpty(creds.PasswordHash)
        || string.IsNullOrEmpty(creds.PasswordSalt);

    public static bool TryParseBasic(string? authHeader, out string user, out string pass)
    {
        user = pass = "";
        if (string.IsNullOrEmpty(authHeader)
            || !authHeader.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
            return false;

        try
        {
            var b64     = authHeader[6..].Trim();
            var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            var idx     = decoded.IndexOf(':');
            if (idx < 0) return false;
            user = decoded[..idx];
            pass = decoded[(idx + 1)..];
            return true;
        }
        catch
        {
            return false;
        }
    }
}
