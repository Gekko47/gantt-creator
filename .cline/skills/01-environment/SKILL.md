---
name: 01-ENVIRONMENT
description: Dev-machine baseline (.NET / Office / Visual Studio / Excel-DNA / VS Code versions), first-run installation checklist, common Office-host failures, and the offline acceptance check. Trigger on environment, setup, install, or version questions.
---

Dev-machine baseline, first-run install checklist, and the offline
acceptance check for this repo.

Do not get wrong:
- Record actual installed versions in the work item and CI output;
  never silently develop against a different baseline than the one
  recorded in the baseline table of `docs/01-ENVIRONMENT.md`.
- Trust the Windows build number, not the `ProductName` registry
  string, which can misreport "Windows 10 Pro" on a Windows 11 host.
- The IDE in use is Visual Studio 2026 Community (see ADR 0003), not
  a generic "VS 2022 17.14+" — check the ADR before assuming version
  parity with the baseline table in `docs/01-ENVIRONMENT.md`.

---

## Tools

- `run_commands` / `pwsh_run` — run `dotnet --info`, `(Get-CimInstance Win32_OperatingSystem).BuildNumber`, and other environment probes.
- `read_files` — read `global.json`, `Directory.Packages.props`, `tool-versions.psd1` to verify pinned versions.
- `dotnet_build` / `dotnet_test` — run the offline acceptance check; build and test without Office.
- `dotnet_packages` — audit NuGet references for outdated/vulnerable versions against the baseline.
- `read_files` on `docs/adr/` — check ADRs before assuming version parity (e.g. ADR 0003 for VS version).

---

## Where to read more

- Full reference: `docs/01-ENVIRONMENT.md` (this is the canonical source; there is no separate copy under .cline/skills/)
