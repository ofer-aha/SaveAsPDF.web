using System.Collections.Concurrent;

/// <summary>
/// Per-IP lockout for the admin login endpoint.
///
/// PBKDF2 makes each guess expensive for us as well as the attacker, so without a
/// throttle a scripted password grind is both feasible and a denial-of-service on
/// the server's CPU. After <see cref="MaxAttempts"/> consecutive failures an IP is
/// locked out for <see cref="LockoutPeriod"/>; a successful login clears the counter.
///
/// In-memory and per-process, matching SessionService. That is adequate here (single
/// instance, LAN-only) but means a service restart clears lockouts.
/// </summary>
public static class LoginThrottle
{
    private const int MaxAttempts = 8;
    private static readonly TimeSpan LockoutPeriod = TimeSpan.FromMinutes(15);
    private static readonly TimeSpan FailureWindow = TimeSpan.FromMinutes(15);

    private sealed class Entry
    {
        public int      Failures    { get; set; }
        public DateTime LastFailure { get; set; }
        public DateTime LockedUntil { get; set; }
    }

    private static readonly ConcurrentDictionary<string, Entry> _entries = new();

    public static bool IsLockedOut(string ip, out TimeSpan retryAfter)
    {
        retryAfter = TimeSpan.Zero;
        if (!_entries.TryGetValue(ip, out var e)) return false;

        if (e.LockedUntil > DateTime.UtcNow)
        {
            retryAfter = e.LockedUntil - DateTime.UtcNow;
            return true;
        }

        // Stale failures decay so a slow trickle of typos never accumulates a lockout.
        if (DateTime.UtcNow - e.LastFailure > FailureWindow)
            _entries.TryRemove(ip, out _);

        return false;
    }

    public static void RecordFailure(string ip)
    {
        Purge();
        var e = _entries.GetOrAdd(ip, _ => new Entry());
        lock (e)
        {
            if (DateTime.UtcNow - e.LastFailure > FailureWindow) e.Failures = 0;
            e.Failures++;
            e.LastFailure = DateTime.UtcNow;
            if (e.Failures >= MaxAttempts)
            {
                e.LockedUntil = DateTime.UtcNow.Add(LockoutPeriod);
                e.Failures    = 0;
                Console.Error.WriteLine(
                    $"[LoginThrottle] {ip} locked out until {e.LockedUntil:HH:mm:ss} UTC " +
                    $"after {MaxAttempts} failed admin logins.");
            }
        }
    }

    public static void RecordSuccess(string ip) => _entries.TryRemove(ip, out _);

    private static void Purge()
    {
        if (_entries.Count < 256) return;
        var now = DateTime.UtcNow;
        foreach (var kv in _entries)
        {
            if (kv.Value.LockedUntil < now && now - kv.Value.LastFailure > FailureWindow)
                _entries.TryRemove(kv.Key, out _);
        }
    }
}
