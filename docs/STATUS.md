# Status

> Short factual handoff for human reviewers and for restarting a Cline session.
> Update after every commit. Do not paste chat transcripts.

## Active work item

- **ID**: R0.2 (and the verify-script piece of R0.6)
- **Title**: Solution, projects, central packages, architecture test, verify scripts
- **Branch**: `stage-inspect`
- **Started**: 2026-09-04
- **Outcome**: in progress — two commits landed; the .NET 10 Release build is green, the architecture test is live, and `verify-quick.ps1` passes end-to-end in ~9 s.
- **Evidence**: commits `aa0815b` (R0.1) and `900f78c` (R0.2 + verify scripts). Latest run: `pwsh ./scripts/verify-quick.ps1` → PASS in 8.8 s.
- **Next safe action**: R0.3 — produce `packages.lock.json`, add a deliberate two-run restore test, record the actual Office build into the work item evidence, and write the work item file `docs/work-items/R0.3-central-packages-and-lockfile.md`.

## Recently completed

- **R0.1** — Governance kit and repo scaffolding. Kit installed at `docs/0N-*.md` and `.clinerules/0N-*.md`; `AGENTS.md`, `docs/DECISIONS.md`, `docs/KNOWN-LIMITATIONS.md`, `docs/REVIEW-CHECKLIST.md`, `docs/work-items/TEMPLATE.md`, the first work item, four ADRs, and `scripts/check-md-links.ps1` committed. Commit `aa0815b`.
- **R0.2 + R0.6 (verify scripts)** — Solution (`GanttCreator.slnx`), six production projects (Core on `net10.0`; Raster, Office, AddIn on `net10.0-windows`), six test projects, central package management with `Directory.Packages.props` (ExcelDna.AddIn 1.9.0, ExcelDna.Integration 1.9.0, ExcelDna.Interop 16.0.0 per ADR-0004), `Directory.Build.props` enforcing nullable + warnings-as-errors + analyzers + deterministic, `global.json` pinning .NET SDK 10.0.400, the architecture test that prevents `GanttCreator.Core` from referencing Office/Excel-DNA/SkiaSharp/clipboard, and the three verify scripts. Commit `900f78c`.

## Environment (recorded once, then referenced)

- **Host OS**: Windows 10 Pro 2009, x64
- **.NET SDK**: 10.0.400
- **.NET runtimes**: 10.0.11, 8.0.30 (Microsoft.NETCore.App, Microsoft.AspNetCore.App, Microsoft.WindowsDesktop.App)
- **Visual Studio**: 18 (Community), installed at `C:\Program Files\Microsoft Visual Studio\18\Community`
- **Office**: Microsoft 365 x64 retail, 16.0.20326.20112, current channel, en-US
- **PowerShell**: 7.4.19
- **Excel-DNA target**: 1.9.0
- **Self-hosted runner policy**: option 1 — a real licensed Microsoft 365 x64 install on this machine; the Office-integration suite is excluded from `verify.ps1` and runs via `scripts/verify-office.ps1` only at phase exit and release.

## Known limitations

- **L1** — Windows 10 (not Windows 11) is the actual host. Excel-DNA 1.9 and Office 16.0.x are supported; recorded in the environment table above.
- **L2** — `.vs/` (Visual Studio user options) is present locally and is already covered by `.gitignore`.
- **L3** — The optional `03-ROADMAP.md` skill originally referenced a `03` gap; the renumbered `.clinerules/03-ROADMAP.md` now closes that gap and the `docs/03-ROADMAP.md` mirror is committed.

## Open questions for human

- _none yet_
