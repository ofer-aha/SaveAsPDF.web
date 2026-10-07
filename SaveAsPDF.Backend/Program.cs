//var builder = WebApplication.CreateBuilder(args);
var options = new WebApplicationOptions { Args = args, ContentRootPath = AppContext.BaseDirectory };
var builder = WebApplication.CreateBuilder(options);

// Emails with inline images (base64-embedded) can produce large JSON payloads.
builder.WebHost.ConfigureKestrel(k => k.Limits.MaxRequestBodySize = 100_000_000); // 100 MB

builder.Services.AddControllers()
    .AddJsonOptions(opts => {
        opts.JsonSerializerOptions.PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase;
        opts.JsonSerializerOptions.MaxDepth = 64;
    });

builder.Services.AddSingleton<SettingsService>();
builder.Services.AddSingleton<ProjectDataService>();
builder.Services.AddSingleton<SessionService>();
builder.Services.AddSingleton<LogService>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("SaveAsPDFCors", policy =>
    {
        policy
            .WithOrigins("https://localhost:3000")
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

var app = builder.Build();

// Apply the admin-selected Chromium build (if any) to the PDF engine on startup.
try
{
    var startupSettings = app.Services.GetRequiredService<SettingsService>().Load();
    PdfService.SetPreferredBuild(startupSettings.ChromiumBuild);
}
catch { /* fall back to the default pinned Chromium build */ }

// ---------------------------------------------------------------
// Auth middleware — protects /api/settings and /api/admin/* (except session endpoint).
// The admin HTML (/admin/*) is intentionally PUBLIC — the page itself shows a login
// form; all data APIs require a session token issued by POST /api/admin/session.
//
// Accepts: X-Admin-Token header (10-minute sliding session token).
// ---------------------------------------------------------------
app.Use(async (ctx, next) =>
{
    var path = ctx.Request.Path.Value ?? "";

    // MVC route matching discards an empty trailing segment, so
    // "/api/project/1000/folders/" reaches exactly the same action as
    // "/api/project/1000/folders". Every test below is a string comparison, so
    // without this normalisation one extra character walks straight past the
    // DELETE gate.
    if (path.Length > 1) path = path.TrimEnd('/');

    bool needsAuth =
        path.StartsWith("/api/settings", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/api/logs",     StringComparison.OrdinalIgnoreCase) ||
        // Opens Explorer on the server console via Process.Start - admin only.
        path.StartsWith("/api/open-folder", StringComparison.OrdinalIgnoreCase) ||
        (path.StartsWith("/api/admin",   StringComparison.OrdinalIgnoreCase) &&
         !path.Equals("/api/admin/session",        StringComparison.OrdinalIgnoreCase) &&
         !path.StartsWith("/api/admin/session/",   StringComparison.OrdinalIgnoreCase));

    // Destructive folder operations require an admin session. Creating and renaming
    // stay open because the taskpane needs them and they are contained to a single
    // project; recursive delete is a different risk class - an unauthenticated LAN
    // host should not be able to remove project directories.
    if (!needsAuth
        && HttpMethods.IsDelete(ctx.Request.Method)
        && path.StartsWith("/api/project/", StringComparison.OrdinalIgnoreCase)
        && path.EndsWith("/folders", StringComparison.OrdinalIgnoreCase))
    {
        needsAuth = true;
    }

    // SSE stream: EventSource can't send custom headers, so also accept ?token= query param
    if (needsAuth && path.StartsWith("/api/logs/stream", StringComparison.OrdinalIgnoreCase))
    {
        var qToken = ctx.Request.Query["token"].FirstOrDefault();
        if (!string.IsNullOrWhiteSpace(qToken))
        {
            var sessions2 = ctx.RequestServices.GetRequiredService<SessionService>();
            if (sessions2.Validate(qToken)) { await next(); return; }
            ctx.Response.StatusCode = 401;
            await ctx.Response.WriteAsync("Unauthorized");
            return;
        }
    }

    if (!needsAuth) { await next(); return; }

    var sessions = ctx.RequestServices.GetRequiredService<SessionService>();
    var token    = ctx.Request.Headers["X-Admin-Token"].FirstOrDefault();
    if (sessions.Validate(token)) { await next(); return; }

    ctx.Response.StatusCode = 401;
    await ctx.Response.WriteAsync("Unauthorized — session token required");
});

// Admin web UI lives at /admin (served from wwwroot/admin/index.html)
app.UseDefaultFiles();
// Disable caching for all static files so add-in updates are picked up immediately.
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers["Cache-Control"] = "no-cache, no-store, must-revalidate";
        ctx.Context.Response.Headers["Pragma"]        = "no-cache";
        ctx.Context.Response.Headers["Expires"]       = "0";
    }
});

app.UseRouting();

app.UseCors("SaveAsPDFCors");

app.MapControllers();

// Session management routes
// POST /api/admin/session — public endpoint; validates JSON credentials, returns a sliding token
app.MapPost("/api/admin/session", async (HttpContext ctx, SettingsService settings, SessionService sessions) =>
{
    // Throttle by source IP. The admin password is the only thing standing between a
    // LAN host and the settings API, and PBKDF2 alone does not stop an attacker
    // grinding a weak password over the network.
    var ip = ctx.Connection.RemoteIpAddress?.ToString() ?? "unknown";
    if (LoginThrottle.IsLockedOut(ip, out var retryAfter))
    {
        ctx.Response.Headers["Retry-After"] = ((int)retryAfter.TotalSeconds).ToString();
        return Results.Json(
            new { error = $"יותר מדי נסיונות. נסה שוב בעוד {(int)retryAfter.TotalSeconds} שניות." },
            statusCode: 429);
    }

    LoginRequest? req = null;
    try { req = await ctx.Request.ReadFromJsonAsync<LoginRequest>(); } catch { }
    if (req == null || string.IsNullOrWhiteSpace(req.Username))
        return Results.BadRequest(new { error = "Username and password required" });
    var creds = settings.Load().Admin;

    // Until a real password is set, admin/admin is accepted only from the server
    // itself. Otherwise a fresh install - or one whose settings.json went missing -
    // is fully administrable by any host on the LAN using a documented default.
    // Escape hatch if you are ever locked out: open https://localhost:5176/admin
    // while signed in on MG01 and set a password there.
    if (AdminAuthService.IsUsingDefaults(creds)
        && !(ctx.Connection.RemoteIpAddress?.Equals(System.Net.IPAddress.Loopback) == true
             || ctx.Connection.RemoteIpAddress?.Equals(System.Net.IPAddress.IPv6Loopback) == true))
    {
        LoginThrottle.RecordFailure(ip);
        return Results.Json(new
        {
            error = "לא הוגדרה סיסמת מנהל. יש להתחבר מהשרת עצמו (localhost) ולהגדיר סיסמה."
        }, statusCode: 403);
    }

    if (!AdminAuthService.Verify(creds, req.Username, req.Password ?? ""))
    {
        LoginThrottle.RecordFailure(ip);
        return Results.Json(new { error = "שם משתמש או סיסמה שגויים" }, statusCode: 401);
    }
    LoginThrottle.RecordSuccess(ip);
    return Results.Ok(new { token = sessions.Create(), expiresInMinutes = 10 });
});

// DELETE /api/admin/session — revoke token on explicit logout
app.MapDelete("/api/admin/session", (HttpContext ctx, SessionService sessions) =>
{
    sessions.Remove(ctx.Request.Headers["X-Admin-Token"].FirstOrDefault());
    return Results.Ok();
});

// POST /api/admin/session/revoke — called by navigator.sendBeacon on tab close (no auth required;
// the token itself is the credential and it expires naturally if this call is dropped)
app.MapPost("/api/admin/session/revoke", (HttpContext ctx, SessionService sessions) =>
{
    var token = ctx.Request.Query["t"].FirstOrDefault();
    sessions.Remove(token);
    return Results.Ok();
});

// Friendly redirects
app.MapGet("/",            ctx => { ctx.Response.Redirect("/admin/index.html");      return Task.CompletedTask; });
app.MapGet("/admin",       ctx => { ctx.Response.Redirect("/admin/index.html");      return Task.CompletedTask; });
app.MapGet("/help",        ctx => { ctx.Response.Redirect("/help/index.html");       return Task.CompletedTask; });
app.MapGet("/help/admin",  ctx => { ctx.Response.Redirect("/help/admin/index.html"); return Task.CompletedTask; });

app.Run();

record LoginRequest(string? Username, string? Password);
