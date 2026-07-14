# Security Policy

## Supported Versions

Security fixes are applied to the latest `4.x` release line. Older major
versions are not maintained.

| Version | Supported |
| ------- | --------- |
| 4.x     | ✅        |
| < 4.0   | ❌        |

## Reporting a Vulnerability

Please **do not** open a public issue for security problems.

Report privately via GitHub's [security advisory form](https://github.com/NeomaVerwaltung/PptxGenJS/security/advisories/new).
Include a description, affected version, and a reproduction if possible.

We aim to acknowledge reports within 5 business days and to ship a fix or
mitigation for confirmed high-severity issues as a priority.

## Scope

This library generates `.pptx` files from untrusted input (text, images,
data). Treat generated output as you would any user-supplied file. Report
any case where crafted input can:

- read or write files outside the intended output path,
- cause the process to hang or exhaust memory, or
- inject unexpected content into the generated document.

Build-time (`devDependencies`) advisories that do not affect the published
`dist/` output are tracked but are lower priority than runtime issues.
