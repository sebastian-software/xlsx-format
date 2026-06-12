# Security Policy

## Supported Versions

Security fixes are provided for the current major release line.

| Version | Supported |
| ------- | --------- |
| 2.x     | Yes       |
| < 2.0   | No        |

## Reporting a Vulnerability

Please do not open a public issue for an unpatched vulnerability.

Use GitHub's private vulnerability reporting flow for this repository. If that flow is unavailable, contact the maintainer through GitHub and request a private reporting channel before sharing exploit details.

Include enough detail to reproduce and assess the issue:

- affected version or commit,
- input file or minimal reproduction,
- expected and actual behavior,
- impact and any known exploitation requirements,
- whether the report is already disclosed elsewhere.

## Response Expectations

The maintainers aim to acknowledge reports within 3 business days and provide an initial assessment or follow-up questions within 7 business days. Fix timelines depend on severity, reproducibility, and release risk.

Coordinated disclosure is preferred. Please allow time for a fix and release before publishing details that would make exploitation straightforward.

## Scope

Reports are especially useful for:

- parsing untrusted XLSX, XLSM, CSV, TSV, or HTML input,
- ZIP, XML, shared-string, worksheet, and relationship parsing,
- prototype-pollution or object-key handling bugs,
- denial-of-service vectors such as excessive memory, CPU, or decompression work,
- CSV/formula injection or unsafe HTML export behavior,
- package export, type-resolution, or build-artifact integrity issues.

xlsx-format does not execute workbook formulas, VBA macros, external entities, or embedded scripts. Bugs that require a separate spreadsheet application to execute generated output may still be security-relevant when xlsx-format is used to export user-controlled data.
