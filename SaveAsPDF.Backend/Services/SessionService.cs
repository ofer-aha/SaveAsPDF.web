using System.Collections.Concurrent;
using System.Security.Cryptography;

/// <summary>
/// In-memory admin session store.
/// Each token is a random hex string with a sliding 10-minute TTL.
/// sessionStorage on the client ensures tokens are discarded when the browser tab is closed.
/// </summary>
public class SessionService
{
    private readonly ConcurrentDictionary<string, DateTime> _sessions = new();
    private static readonly TimeSpan Ttl = TimeSpan.FromMinutes(10);

    /// <summary>Creates a new session token and returns it.</summary>
    public string Create()
    {
        Purge();
        var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(32)).ToLowerInvariant();
        _sessions[token] = DateTime.UtcNow.Add(Ttl);
        return token;
    }

    /// <summary>
    /// Returns true and slides the TTL if the token is valid and not expired.
    /// Returns false (and removes the token) if expired or unknown.
    /// </summary>
    public bool Validate(string? token)
    {
        if (string.IsNullOrEmpty(token)) return false;
        if (!_sessions.TryGetValue(token, out var exp)) return false;
        if (exp < DateTime.UtcNow) { _sessions.TryRemove(token, out _); return false; }
        _sessions[token] = DateTime.UtcNow.Add(Ttl); // sliding window
        return true;
    }

    /// <summary>Explicitly invalidates a token (logout).</summary>
    public void Remove(string? token)
    {
        if (!string.IsNullOrEmpty(token)) _sessions.TryRemove(token, out _);
    }

    private void Purge()
    {
        var now  = DateTime.UtcNow;
        var dead = _sessions.Where(kv => kv.Value < now).Select(kv => kv.Key).ToList();
        foreach (var key in dead) _sessions.TryRemove(key, out _);
    }
}
