# SaveAsPDF.Web

SaveAsPDF.Web is an Outlook add-in + ASP.NET backend solution for saving selected Outlook messages to PDF and storing them in project folders, including metadata stamping and attachment handling.

## Table of Contents
- [Overview](#overview)
- [Architecture](#architecture)
- [Repository Layout](#repository-layout)
- [Prerequisites](#prerequisites)
- [Local Development](#local-development)
- [Build and Run](#build-and-run)
- [Manifest and Sideloading](#manifest-and-sideloading)
- [Backend API and Security Model](#backend-api-and-security-model)
- [Versioning](#versioning)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)

## Overview

The repository contains:

- **Outlook web add-in frontend** (Office.js, Webpack, task pane + command surface)
- **.NET backend service** (`SaveAsPDF.Backend`) that serves static assets/admin UI and API endpoints
- **Manifest files** for Outlook add-in deployment/testing (`manifest.xml`, `manifest.json`)

Primary flow:

1. User opens an email in Outlook.
2. User runs SaveAsPDF command from the ribbon.
3. Task pane/command code gathers message context.
4. Backend receives content and options, renders/stores output, and manages admin/config/session behavior.

## Architecture

### Frontend (Outlook Add-in)

- Built with `office-addin-*` tooling and Webpack.
- Entrypoints include:
  - `taskpane.html` / `taskpane.js`
  - `commands.html` / `commands.js`
- Served over HTTPS for Office add-in runtime.

### Backend (`SaveAsPDF.Backend`)

- ASP.NET Core Web app targeting `net10.0-windows`.
- Hosts:
  - API controllers
  - Admin web UI under `/admin`
  - Help pages under `/help`
  - Static files (`wwwroot`)
- Uses `PuppeteerSharp` for HTML→PDF rendering.

## Repository Layout

```text
.
├─ README.md
├─ SECURITY.md
├─ package.json
├─ manifest.xml
├─ manifest.json
├─ webpack.config.js
├─ src/
│  ├─ commands/
│  ├─ contactpicker/
│  └─ taskpane/
├─ assets/
├─ dist/
└─ SaveAsPDF.Backend/
   ├─ Controllers/
   ├─ Models/
   ├─ Services/
   ├─ wwwroot/
   ├─ Program.cs
   └─ SaveAsPDF.Backend.csproj
```

## Prerequisites

### Required

- **Node.js** (LTS recommended)
- **npm**
- **Microsoft 365 Outlook** (desktop/web test target)
- **.NET SDK 10** (preview/compatible with project target)
- **Windows** environment (project targets `net10.0-windows`)

### Recommended

- Office Add-in tooling familiarity (`office-addin-debugging`, manifest sideloading)
- A local HTTPS development environment

## Local Development

Install dependencies:

```bash
npm install
```

Common scripts (from `package.json`):

- `npm run build` – production webpack bundle
- `npm run build:dev` – development bundle
- `npm run dev-server` – webpack dev server
- `npm run start` – start Office add-in debugging with `manifest.json`
- `npm run stop` – stop Office add-in debugging
- `npm run validate` – validate add-in manifest
- `npm run lint` / `npm run lint:fix`

## Build and Run

### Frontend bundle

```bash
npm run build
```

Windows helper:

```bat
_build.bat
```

### Backend

From repository root:

```bash
dotnet build SaveAsPDF.Backend/SaveAsPDF.Backend.csproj
```

Run backend:

```bash
dotnet run --project SaveAsPDF.Backend/SaveAsPDF.Backend.csproj
```

Backend behavior highlights:

- Configures request body limit to 100MB (for large inline-image emails)
- Serves static files with no-cache headers
- Enables CORS policy for `https://localhost:3000`

## Manifest and Sideloading

This repo contains both XML and JSON manifests:

- `manifest.xml`
- `manifest.json`
- built/runtime copies under `dist/` and backend `wwwroot/`

### Dev endpoints

- JSON manifest uses `https://localhost:3000/...` for local debug runtime pages.
- XML manifest includes environment-specific URLs (e.g., `https://MG01:5176/...`).

### Typical development flow

1. Start frontend dev server (`npm run dev-server` or `npm run start`).
2. Validate manifest (`npm run validate`).
3. Sideload manifest into Outlook test client.
4. Open an email and launch the task pane command.

> Keep hostnames, ports, and HTTPS certificates aligned between manifest entries and running services.

## Backend API and Security Model

Based on `Program.cs`:

- Admin UI pages are public static pages (`/admin/*`) that provide login UX.
- Protected APIs include:
  - `/api/settings`
  - `/api/logs`
  - most `/api/admin/*` endpoints (except session creation/revoke paths)
- Auth uses short-lived session tokens via header:
  - `X-Admin-Token`
- Session endpoint:
  - `POST /api/admin/session` (returns token)
- Token revoke endpoints:
  - `DELETE /api/admin/session`
  - `POST /api/admin/session/revoke`

For SSE log streaming where custom headers are unavailable, token query param is accepted specifically on logs stream route.

## Versioning

Single source of truth is **`package.json` `version`**.

`SaveAsPDF.Backend.csproj` reads this value at build time and applies it to:

- Assembly version
- File version
- Informational version

To bump version, edit `package.json` only.

## Troubleshooting

### Add-in does not load in Outlook

- Ensure HTTPS endpoint in manifest is reachable.
- Ensure certificate is trusted for localhost/dev hostname.
- Re-run `npm run validate` and fix manifest errors.

### CORS/API failures from task pane

- Confirm frontend origin matches backend CORS allowlist (`https://localhost:3000`).
- Check backend is running and accessible.

### PDF generation issues

- Verify backend can start/access Chromium via `PuppeteerSharp`.
- Check server logs for rendering exceptions.

### Version mismatch between frontend/backend

- Confirm `package.json` version was updated before backend build.
- Rebuild backend after changing package version.

## Contributing

1. Create a feature branch.
2. Keep frontend and backend changes coherent.
3. Validate manifests and lint before PR.
4. Include clear reproduction/verification steps in pull requests.
