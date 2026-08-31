# Security policy

## Security status

`hwp2yaml` parses complex document and archive formats. The current public baseline is maintained on a best-effort basis, but no published version is certified as safe for arbitrary hostile documents until the bounded-parser, external-process, dependency and synthetic adversarial gates in `AGENTS.md` pass.

Do not process sensitive or untrusted documents in a privileged account, shared host or network-enabled production environment solely on the basis of the current package name or README claims.

## Reporting a vulnerability

Private vulnerability reporting is enabled for this repository. Use [GitHub's private vulnerability report form](https://github.com/seunghyuoffice-design/hwp2yaml/security/advisories/new).

- Do not attach real HWP/HWPX/PDF files, customer data, account data, credentials, absolute personal paths or full converted output.
- Prefer a minimal generator, synthetic bytes, a redacted stack trace and affected commit/version.

Please include:

- affected commit, package version and platform;
- whether the issue is parser resource exhaustion, path/privacy disclosure, malformed-input handling, external process/network behavior, dependency/provenance or output integrity;
- the smallest synthetic reproduction and expected versus observed result;
- whether output or temporary artifacts were created after failure.

## Disclosure and fixes

Security fixes must preserve provenance and privacy boundaries, add a positive or negative regression test as appropriate, and pass exact-head hosted CI plus final fresh-clone verification before release. A failing test, quota or scan may not be removed merely to close a report.
