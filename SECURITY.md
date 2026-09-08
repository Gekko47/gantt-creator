# Security policy

## Supported versions

The project is **pre-release**: no version has shipped yet, so there is no supported release line. Security fixes are applied to the in-development `stage-inspect` branch and released with the first stable release.

## Reporting a vulnerability

To report a security vulnerability in Gantt Creator:

1. **Do not** open a public GitHub issue for the vulnerability itself.
2. Contact the maintainer directly via the repository's GitHub "Security" tab (advisory) or through the contact listed in the project profile.
3. Include a minimal reproduction, the affected version, and the potential impact.

The maintainer acknowledges reports within 5 business days and aims to ship a fix or mitigation in the next patch release.

## Product security properties

- **Offline operation**: Core functionality requires no network access. There is no telemetry, no cloud API, no online licence check, and no hidden network fallback.
- **No customer data in tests**: Tests use synthetic construction data only. No customer schedule data is committed.
- **Privacy redaction**: The local rolling log redacts emails, file paths, dates, GUIDs, and long hex tokens before writing.
- **Dependency policy**: Packages are centrally managed with locked restore (`Directory.Packages.props` + `packages.lock.json`). Vulnerabilities are scanned via `dotnet list package --vulnerable --include-transitive` in the verify gates.
