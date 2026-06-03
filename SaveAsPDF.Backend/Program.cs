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

    bool needsAuth =
        path.StartsWith("/api/settings", StringComparison.OrdinalIgnoreCase) ||
        (path.StartsWith("/api/admin",   StringComparison.OrdinalIgnoreCase) &&
         !path.Equals("/api/admin/session",        StringComparison.OrdinalIgnoreCase) &&
         !path.StartsWith("/api/admin/session/",   StringComparison.OrdinalIgnoreCase));

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
    LoginRequest? req = null;
    try { req = await ctx.Request.ReadFromJsonAsync<LoginRequest>(); } catch { }
    if (req == null || string.IsNullOrWhiteSpace(req.Username))
        return Results.BadRequest(new { error = "Username and password required" });
    var creds = settings.Load().Admin;
    if (!AdminAuthService.Verify(creds, req.Username, req.Password ?? ""))
        return Results.Json(new { error = "שם משתמש או סיסמה שגויים" }, statusCode: 401);
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
