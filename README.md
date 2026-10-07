# SaveAsPDF.Web

Outlook task-pane add-in that saves e-mails (plus attachments) as stamped PDFs into per-project folders, with an admin console for policy control.

## Architecture

| Part | Tech | Location |
|---|---|---|
| Frontend (taskpane) | JS + webpack, Office.js | `src/taskpane/` |
| Backend | .NET 10 Kestrel, self-hosted | `SaveAsPDF.Backend/` |
| PDF engine | PuppeteerSharp + headless Chromium (RTL-safe) | `SaveAsPDF.Backend/Services/PdfService.cs` |
| Admin console | static HTML served from wwwroot | `SaveAsPDF.Backend/wwwroot/admin/` |
| Help pages (canonical) | static HTML | `publish/wwwroot/help/` (user) · `publish/wwwroot/help/admin/` (admin) |

Runs on server **MG01** as Windows service **`SaveAsPDF`** (via NSSM), serving `https://mg01:5176`.
Deployed app folder: `C:\Apps\SaveAsPDF` (= `Z:\APPS\SaveAsPDF` over the network share).

## Build & deploy

All deploys go through **`C:\dev\SaveAsPDF\Deploy_SaveAsPDF.ps1`** (run elevated on MG01):

```powershell
.\Deploy_SaveAsPDF.ps1 -b    # bump build   (x.y.Z+1)
.\Deploy_SaveAsPDF.ps1 -mi   # bump minor   (x.Y+1.0)
.\Deploy_SaveAsPDF.ps1 -mg   # bump major   (X+1.0.0)
```

The script stops the service, bumps `manifest.xml` + `package.json`, builds webpack + the .NET solution, copies binaries and help pages, restarts, and smoke-tests `/api/info`.
It then commits everything as `Release vX.Y.Z`, tags `vX.Y.Z` and pushes branch + tag to `origin` (normal push, never `--force`; a git failure only warns). `-m "text"` adds to the commit message, `-nogit` skips the git step.

### Backup

`C:\dev\SaveAsPDF\Backup-SaveAsPDF.ps1` (run elevated on MG01) zips what git does not hold - the help pages (`publish/` is git-ignored), `appsettings.Production.json`, `cert.pfx`, and the service's `settings.json` / `logs.json` - to `C:\Apps\AdminCenter\Backups\`, then commits + pushes. `-Preview` shows what it would do, `-NoGit` skips the push.
Keep all `.ps1` files **pure ASCII** (PowerShell 5.1 reads BOM-less files as ANSI).

### Optional IIS hosting

`C:\dev\SaveAsPDF\Install-SaveAsPDF-IIS.ps1` (run elevated on MG01) migrates hosting from the NSSM service to IIS: installs the IIS role + .NET 10 Hosting Bundle, creates the app pool (No Managed Code, **LocalSystem** — keeps existing settings.json/logs.json) and site on the same `https://*:5176`, disables the old service, smoke-tests, and auto-rolls-back on failure. `-Rollback` returns to the service. Caveat: under IIS the Kestrel `SslProtocols` override does not apply (OS Schannel decides) — if the ESET TLS 1.0 downgrade issue is active, stay on the service.

## Server-only configuration (never overwritten by deploys)

`C:\Apps\SaveAsPDF\appsettings.Production.json` — exists **only on the server**, not in this repo:

```json
{
    "Kestrel": {
        "Endpoints": {
            "Https": {
                "Url": "https://0.0.0.0:5176",
                "SslProtocols": [ "Tls", "Tls11", "Tls12" ],
                "Certificate": { "Path": "C:\\Apps\\SaveAsPDF\\cert.pfx", "Password": "***" }
            }
        }
    }
}
```

Two hard-won rules (July 2026 outage):

1. **`SslProtocols` must include `Tls` (1.0)** — corporate ESET intercepts all TLS and its proxy leg to internal servers can negotiate only TLS 1.0. .NET 10 defaults to TLS 1.2+, which makes every intercepted handshake fail (SSPI 0x80090331). Remove TLS 1.0 only after ESET gets a protocol-filtering exclusion for SaveAsPDF.
2. **`cert.pfx` and the client-side trusted cert must be the same certificate.** The self-signed `CN=mg01` cert must be installed as Trusted Root on every user machine; replacing it on the server without updating clients silently breaks the add-in (WebView2 shows Outlook's generic "couldn't start this add-in"). Previous cert is kept as `cert_old_*.pfx`.

## Smoke testing

Use `curl.exe`, **not** `Invoke-WebRequest` (the latter fails through ESET's TLS interception even when the server is healthy):

```
curl.exe -s https://mg01:5176/api/info      # no -k: also validates the trust chain
```

Chrome may cache a cert error after a cert swap — fully restart it.

## Backend-down UX

When the taskpane loads but the backend is unreachable, it shows a full-pane "sleeping server" overlay (Hebrew, SVG cartoon) with a retry button that auto-reloads once `/api/info` answers. Hooks: startup `/api/info` check and `isNetworkError()` in the save handler (`src/taskpane/taskpane.js`).

## Docs

- User guide: `https://mg01:5176/help`
- Admin guide (auth, policy, HTTPS/TLS, troubleshooting): `https://mg01:5176/help/admin`
- Service logs: `C:\Apps\SaveAsPDF\logs\stdout.log` / `stderr.log`; save-event logs persist in `%LOCALAPPDATA%\SaveAsPDF\logs.json`
