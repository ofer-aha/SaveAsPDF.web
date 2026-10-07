# Security Policy

## Supported Versions

Security updates are provided on a best-effort basis for the latest code on the default branch (`master`).

## Reporting a Vulnerability

If you discover a security vulnerability, please report it **privately** and do not open a public issue.

### Preferred contact

- GitHub Security Advisories (private report) for this repository
- Or contact the maintainer directly through a private channel

Please include:

- A clear description of the vulnerability
- Affected component(s) and file paths
- Reproduction steps / proof of concept
- Impact assessment (confidentiality, integrity, availability)
- Suggested remediation (if available)

## Response Expectations

- Initial acknowledgment target: **within 7 days**
- Triage and severity assessment: as soon as possible after acknowledgment
- Fix timeline depends on impact and complexity

## Disclosure Policy

- Please allow reasonable time for investigation and patching before public disclosure.
- Once fixed, coordinated disclosure and release notes are encouraged.

## Security Considerations for This Project

This repository includes an Outlook add-in and backend API. Key areas to review:

- Manifest trust boundaries (allowed domains, HTTPS endpoints)
- Backend admin/session authentication (`X-Admin-Token` and session endpoints)
- CORS configuration and origin restrictions
- PDF rendering pipeline hardening (Chromium/PuppeteerSharp)
- Input validation for email content and attachment metadata
- Logging practices to avoid leaking sensitive message content

## Recommended Hardening Practices

- Keep Node.js, npm dependencies, and NuGet packages up to date.
- Restrict production CORS origins to known hosts only.
- Enforce HTTPS everywhere; avoid mixed-content endpoints.
- Rotate secrets and admin credentials regularly.
- Use least privilege for runtime/service accounts.
- Monitor and alert on repeated authentication failures.

## Out of Scope

Unless explicitly stated otherwise, the following are typically out of scope:

- Denial-of-service requiring unrealistic resources
- Reports based only on outdated/unpatched local setups
- Social engineering or phishing simulations
