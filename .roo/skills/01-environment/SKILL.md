---
name: 01-environment
description: Dev-machine baseline for the Gantt Creator add-in (.NET 10 SDK, .NET/Office/Visual Studio 2026/Excel-DNA 1.9/VS Code versions), first-run installation checklist, Excel-DNA first-run behaviour, common Office-host failures, and the offline acceptance check. Part of the gantt-creator index. Trigger on environment, setup, install, SDK or version questions, or host load failures.
---

# Environment and first run

Mirror of [`.cline/skills/01-environment/SKILL.md`](../../../.cline/skills/01-environment/SKILL.md),
generated from [`docs/01-ENVIRONMENT.md`](../../../docs/01-ENVIRONMENT.md). The canonical doc
wins on any disagreement.

## Index

- Index first: `gantt-creator-contract`. Call `00-ponytail` before changing code.
- Pair with `05-git-quality` for gates and CI parity, and `04-test-strategy` for what the
  offline acceptance check must cover.

## Do not get wrong

- Record actual installed versions in the work item and CI output; never silently develop
  against a different baseline than the one recorded in the baseline table of
  `docs/01-ENVIRONMENT.md`.
- Trust the Windows build number, not the `ProductName` registry string, which can misreport
  "Windows 10 Pro" on a Windows 11 host.
- The IDE in use is Visual Studio 2026 Community (ADR-0003), not a generic "VS 2022 17.14+".
  Check the ADR before assuming version parity with the baseline table.
- `global.json` pins the SDK (`10.0.400`, `rollForward: latestFeature`) and package versions
  are central (`Directory.Packages.props` + committed `packages.lock.json`). Verify, do not
  assume.

## Tools here

`execute_command` (`dotnet --info`, `dotnet --version`, version probes, build and test without
Office) · `read_file` (`global.json`, `Directory.Packages.props`,
`scripts/tool-versions.psd1`, `docs/adr/`) · `search_files` for pin literals shared between
local scripts and CI.
