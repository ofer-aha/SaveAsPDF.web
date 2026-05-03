var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers()
    .AddJsonOptions(opts =>
        opts.JsonSerializerOptions.PropertyNamingPolicy =
            System.Text.Json.JsonNamingPolicy.CamelCase);

builder.Services.AddSingleton<SettingsService>();
builder.Services.AddSingleton<ProjectDataService>();

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
// Basic Auth — protects /admin/*, /api/settings, /api/admin/*.
// Public: /api/info, /api/policy, /api/saveaspdf, /api/project,
// /api/contacts, /api/folders, /help/* (the help pages stay public).
// First-time login: admin/admin (force change in the UI).
// ---------------------------------------------------------------
app.Use(async (ctx, next) =>
{
    var path = ctx.Request.Path.Value ?? "";

    bool needsAuth =
        path.StartsWith("/admin",        StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/api/settings", StringComparison.OrdinalIgnoreCase) ||
        path.StartsWith("/api/admin",    StringComparison.OrdinalIgnoreCase);

    if (!needsAuth) { await next(); return; }

    var settings = ctx.RequestServices.GetRequiredService<SettingsService>();
    var creds    = settings.Load().Admin;

    AdminAuthService.TryParseBasic(ctx.Request.Headers.Authorization, out var u, out var p);
    if (!AdminAuthService.Verify(creds, u, p))
    {
        ctx.Response.StatusCode = 401;
        ctx.Response.Headers.Append("WWW-Authenticate", "Basic realm=\"SaveAsPDF Admin\"");
        await ctx.Response.WriteAsync("Unauthorized");
        return;
    }
    await next();
});

// Admin web UI lives at /admin (served from wwwroot/admin/index.html)
app.UseDefaultFiles();
app.UseStaticFiles();

app.UseRouting();

app.UseCors("SaveAsPDFCors");

app.MapControllers();

// Friendly redirects
app.MapGet("/",            ctx => { ctx.Response.Redirect("/admin/index.html");      return Task.CompletedTask; });
app.MapGet("/admin",       ctx => { ctx.Response.Redirect("/admin/index.html");      return Task.CompletedTask; });
app.MapGet("/help",        ctx => { ctx.Response.Redirect("/help/index.html");       return Task.CompletedTask; });
app.MapGet("/help/admin",  ctx => { ctx.Response.Redirect("/help/admin/index.html"); return Task.CompletedTask; });

app.Run();
