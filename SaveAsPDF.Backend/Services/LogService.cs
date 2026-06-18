using System.Text.Json;
using System.Threading.Channels;

/// <summary>
/// Thread-safe log buffer with optional SSE fan-out, backed by a JSON file so
/// entries survive a service restart or redeploy.
///
/// The in-memory list is the working set (bounded by retention); on every change
/// it is flushed to <c>%LOCALAPPDATA%\SaveAsPDF\logs.json</c>. On construction the
/// file is read back so previous logs are restored on startup.
///
/// Retention (maxEntries, retainDays) is read from settings on every Log() call so
/// changes made in the admin UI take effect immediately.
/// </summary>
public class LogService
{
    private readonly LinkedList<LogEntry> _entries = new();
    private readonly object              _lock     = new();
    private readonly List<Channel<LogEntry>> _subscribers = new();
    private readonly SettingsService     _settings;
    private readonly string              _file;

    private static readonly JsonSerializerOptions _io = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    public LogService(SettingsService settings)
    {
        _settings = settings;

        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SaveAsPDF");
        Directory.CreateDirectory(dir);
        _file = Path.Combine(dir, "logs.json");

        LoadFromDisk();
    }

    // Restore previously persisted entries on startup (newest kept within retention).
    private void LoadFromDisk()
    {
        try
        {
            if (!File.Exists(_file)) return;
            var json = File.ReadAllText(_file);
            var list = JsonSerializer.Deserialize<List<LogEntry>>(json, _io);
            if (list == null) return;

            // Oldest-first in the buffer (GetAll sorts newest-first for display).
            foreach (var e in list.OrderBy(e => e.Timestamp))
                _entries.AddLast(e);

            Trim(_settings.Load().LogSettings);
        }
        catch { /* a corrupt/locked log file must never stop the service from starting */ }
    }

    // Trim the in-memory buffer to the configured retention. Caller holds _lock
    // (or is the single-threaded constructor).
    private void Trim(LogSettings s)
    {
        var maxEntries = Math.Max(1, s.MaxEntries);
        while (_entries.Count > maxEntries)
            _entries.RemoveFirst();

        if (s.RetainDays > 0)
        {
            var cutoff = DateTime.Now.AddDays(-s.RetainDays);
            while (_entries.Count > 0 && _entries.First!.Value.Timestamp < cutoff)
                _entries.RemoveFirst();
        }
    }

    // Persist the current buffer to disk. Caller holds _lock. Best-effort.
    private void Persist()
    {
        try { File.WriteAllText(_file, JsonSerializer.Serialize(_entries, _io)); }
        catch { /* never let a disk error break the save flow */ }
    }

    // ── Write ─────────────────────────────────────────────────────────────────

    public void Log(LogEntry entry)
    {
        var s = _settings.Load().LogSettings;

        lock (_lock)
        {
            _entries.AddLast(entry);
            Trim(s);
            Persist();

            // Push to SSE subscribers (fire-and-forget, drop if channel is full)
            foreach (var ch in _subscribers.ToList())
                ch.Writer.TryWrite(entry);
        }
    }

    // ── Read ──────────────────────────────────────────────────────────────────

    public List<LogEntry> GetAll(
        string?   level  = null,
        string?   search = null,
        DateTime? from   = null,
        DateTime? to     = null)
    {
        lock (_lock)
        {
            IEnumerable<LogEntry> q = _entries;

            if (!string.IsNullOrWhiteSpace(level))
                q = q.Where(e => string.Equals(e.Level, level, StringComparison.OrdinalIgnoreCase));

            if (from.HasValue) q = q.Where(e => e.Timestamp >= from.Value);
            if (to.HasValue)   q = q.Where(e => e.Timestamp <= to.Value);

            if (!string.IsNullOrWhiteSpace(search))
            {
                var s = search.ToLowerInvariant();
                q = q.Where(e =>
                    e.Username.ToLowerInvariant().Contains(s)    ||
                    e.Subject.ToLowerInvariant().Contains(s)     ||
                    e.ProjectId.ToLowerInvariant().Contains(s)   ||
                    e.ProjectName.ToLowerInvariant().Contains(s) ||
                    e.Attachments.Any(a => a.ToLowerInvariant().Contains(s)));
            }

            // Return newest-first so the table shows the latest events at the top
            return q.OrderByDescending(e => e.Timestamp).ToList();
        }
    }

    public int Count() { lock (_lock) return _entries.Count; }

    public void Clear() { lock (_lock) { _entries.Clear(); Persist(); } }

    // ── SSE pub/sub ───────────────────────────────────────────────────────────

    /// <summary>Register a bounded channel that receives every new log entry.</summary>
    public Channel<LogEntry> Subscribe()
    {
        var ch = Channel.CreateBounded<LogEntry>(new BoundedChannelOptions(200)
        {
            FullMode         = BoundedChannelFullMode.DropOldest,
            SingleReader     = true,
            SingleWriter     = false,
            AllowSynchronousContinuations = false
        });
        lock (_lock) { _subscribers.Add(ch); }
        return ch;
    }

    /// <summary>Deregister and complete the channel when the SSE client disconnects.</summary>
    public void Unsubscribe(Channel<LogEntry> ch)
    {
        lock (_lock) { _subscribers.Remove(ch); }
        ch.Writer.TryComplete();
    }
}
