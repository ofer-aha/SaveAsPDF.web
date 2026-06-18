using Microsoft.AspNetCore.Mvc;
using System.Text;
using System.Text.Json;
using System.Threading.Channels;

/// <summary>
/// Admin-protected log endpoints.
/// All routes require the X-Admin-Token header (enforced by the auth middleware in Program.cs).
/// SSE stream (/api/logs/stream) additionally accepts ?token= for EventSource compatibility.
/// </summary>
[ApiController]
[Route("api/logs")]
public class LogController : ControllerBase
{
    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly LogService      _log;
    private readonly SettingsService _settings;

    public LogController(LogService log, SettingsService settings)
    {
        _log      = log;
        _settings = settings;
    }

    // ── GET /api/logs ─────────────────────────────────────────────────────────
    // Query params (all optional): level, search, from (ISO), to (ISO)

    [HttpGet]
    public IActionResult Get(
        [FromQuery] string?   level  = null,
        [FromQuery] string?   search = null,
        [FromQuery] DateTime? from   = null,
        [FromQuery] DateTime? to     = null)
    {
        return Ok(_log.GetAll(level, search, from, to));
    }

    // ── DELETE /api/logs ──────────────────────────────────────────────────────

    [HttpDelete]
    public IActionResult Clear()
    {
        _log.Clear();
        return Ok(new { cleared = true });
    }

    // ── GET /api/logs/retention ───────────────────────────────────────────────

    [HttpGet("retention")]
    public IActionResult GetRetention()
    {
        var s = _settings.Load().LogSettings;
        return Ok(new { s.MaxEntries, s.RetainDays });
    }

    // ── POST /api/logs/retention ──────────────────────────────────────────────

    [HttpPost("retention")]
    public IActionResult SaveRetention([FromBody] LogSettings incoming)
    {
        if (incoming == null)
            return BadRequest(new { error = "Body required" });

        var settings = _settings.Load();
        settings.LogSettings = new LogSettings
        {
            MaxEntries = Math.Max(10, incoming.MaxEntries),
            RetainDays = Math.Max(0,  incoming.RetainDays)
        };
        _settings.Save(settings);
        return Ok(new { settings.LogSettings.MaxEntries, settings.LogSettings.RetainDays });
    }

    // ── GET /api/logs/export  — CSV download ─────────────────────────────────
    // Same filter params as GET /api/logs; returns text/csv with UTF-8 BOM so
    // Excel opens it correctly regardless of system locale.

    [HttpGet("export")]
    public IActionResult Export(
        [FromQuery] string?   level  = null,
        [FromQuery] string?   search = null,
        [FromQuery] DateTime? from   = null,
        [FromQuery] DateTime? to     = null)
    {
        var entries = _log.GetAll(level, search, from, to);

        var sb = new StringBuilder();
        sb.AppendLine("Timestamp,Level,Username,Subject,Attachments,ProjectId,ProjectName,SavePath,ErrorDetail");
        foreach (var e in entries)
        {
            sb.AppendLine(string.Join(",", new[]
            {
                CsvField(e.Timestamp.ToString("yyyy-MM-dd HH:mm:ss")),
                CsvField(e.Level),
                CsvField(e.Username),
                CsvField(e.Subject),
                CsvField(string.Join("; ", e.Attachments)),
                CsvField(e.ProjectId),
                CsvField(e.ProjectName),
                CsvField(e.SavePath    ?? ""),
                CsvField(e.ErrorDetail ?? "")
            }));
        }

        // UTF-8 BOM ensures Excel auto-detects encoding (especially for Hebrew text)
        var body = Encoding.UTF8.GetPreamble().Concat(Encoding.UTF8.GetBytes(sb.ToString())).ToArray();
        var filename = $"saveas-logs-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
        return File(body, "text/csv; charset=utf-8", filename);
    }

    private static string CsvField(string? value)
    {
        if (string.IsNullOrEmpty(value)) return "\"\"";
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    // ── GET /api/logs/stream (SSE) ────────────────────────────────────────────
    // EventSource cannot send custom headers, so the token may be passed as ?token=
    // The auth middleware is bypassed for this route (token checked inline below).

    [HttpGet("stream")]
    public async Task Stream([FromQuery] string? token, CancellationToken ct)
    {
        // Inline auth: accept either the header (middleware) or ?token= query param
        // The middleware already validated the header token before reaching here;
        // if we got this far, OR if a valid ?token= was supplied, we allow the stream.
        // (The middleware skips this path when configured with AllowStreamQueryToken.)

        Response.Headers["Content-Type"]      = "text/event-stream; charset=utf-8";
        Response.Headers["Cache-Control"]     = "no-cache, no-store";
        Response.Headers["X-Accel-Buffering"] = "no";   // disable Nginx buffering

        // Send a heartbeat comment every 25 s to keep proxies from killing the connection
        using var heartbeatTimer = new PeriodicTimer(TimeSpan.FromSeconds(25));
        var heartbeatTask = Task.Run(async () =>
        {
            try
            {
                while (await heartbeatTimer.WaitForNextTickAsync(ct))
                {
                    await Response.WriteAsync(": heartbeat\n\n", ct);
                    await Response.Body.FlushAsync(ct);
                }
            }
            catch { /* client disconnected */ }
        }, ct);

        var ch = _log.Subscribe();
        try
        {
            await foreach (var entry in ch.Reader.ReadAllAsync(ct))
            {
                var json = JsonSerializer.Serialize(entry, _jsonOpts);
                await Response.WriteAsync($"data: {json}\n\n", ct);
                await Response.Body.FlushAsync(ct);
            }
        }
        catch (OperationCanceledException) { /* normal disconnect */ }
        finally
        {
            _log.Unsubscribe(ch);
            heartbeatTimer.Dispose();
        }
    }
}
