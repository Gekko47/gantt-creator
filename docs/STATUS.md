# Status

> Short factual handoff for human reviewers and for restarting a Cline session.
> Update after every commit. Do not paste chat transcripts.

## Active work item

- **ID**: R0.x — three-view rule discipline complete; ready to return to the original Phase 0 sequence.
- **Title**: Cline skills in the correct locations, with sync script and drift gate wired into every commit
- **Branch**: `stage-inspect`
- **Started**: 2026-09-04
- **Outcome**: done. The kit now has `docs/0N-*.md` as the canonical source, `.clinerules/0N-*.md` as the always-on view, and `.cline/skills/0N-name/{SKILL.md,references.md}` as the on-demand view. `scripts/sync-cline-skills.ps1` regenerates the second and third from the first; `scripts/check-cline-skills.ps1` is the drift gate and is wired into `verify-quick.ps1` and `verify.ps1`. The actual host is recorded as Windows 11 25H2 (build 26200) per ADR-0005.
- **Evidence**: commits `aa0815b` (R0.1), `900f78c` (R0.2 + verify scripts), `f232e57` (status), `1645c1e` (Windows 11 25H2 + ADR-0005), `8650717` (skill tree + sync + drift gate), `d543e60` (gate fix: working-tree check, not post-sync). `pwsh ./scripts/verify-quick.ps1` → PASS in ~14 s.
- **Next safe action**: R0.5 — replace `UnitTest1.cs` placeholders in each test project with a non-trivial sample test; apply `[Trait("Category","OfficeIntegration")]` to `GanttCreator.Office.IntegrationTests`; add `.vscode/tasks.json` mirroring `verify-quick.ps1`.

## Recently completed

- **R0.1** — Governance kit and repo scaffolding. Kit installed at `docs/0N-*.md` and `.clinerules/0N-*.md`; `AGENTS.md`, `docs/DECISIONS.md`, `docs/KNOWN-LIMITATIONS.md`, `docs/REVIEW-CHECKLIST.md`, `docs/work-items/TEMPLATE.md`, the first work item, four ADRs, and `scripts/check-md-links.ps1` committed. Commit `aa0815b`.
- **R0.2 + R0.6 (verify scripts)** — Solution (`GanttCreator.slnx`), six production projects (Core on `net10.0`; Raster, Office, AddIn on `net10.0-windows`), six test projects, central package management with `Directory.Packages.props` (ExcelDna.AddIn 1.9.0, ExcelDna.Integration 1.9.0, ExcelDna.Interop 16.0.0 per ADR-0004), `Directory.Build.props` enforcing nullable + warnings-as-errors + analyzers + deterministic, `global.json` pinning .NET SDK 10.0.400, the architecture test that prevents `GanttCreator.Core` from referencing Office/Excel-DNA/SkiaSharp/clipboard, and the three verify scripts. Commit `900f78c`.
- **ADR-0005** — Document the actual host as Windows 11 25H2 (build 26200). The `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion` `ProductName` string may report "Windows 10 Pro" on some configurations; the build number is the source of truth. Commit `1645c1e`.
- **Three-view rule discipline** — `scripts/sync-cline-skills.ps1` regenerates `.clinerules/0N-*.md` and `.cline/skills/0N-name/{SKILL.md,references.md}` from `docs/0N-*.md`; `scripts/check-cline-skills.ps1` is the drift gate; both `verify-quick.ps1` and `verify.ps1` run the gate. `docs/clinerules/SYNC.md` documents the discipline. Commit `8650717`.
- **Drift gate fix** — `check-cline-skills.ps1` was checking post-sync state, which silently overwrote any hand-edited rule. It now checks the working tree directly and fails if any of the three views has uncommitted changes. Commit `d543e60`.
- **R0.3 — Central packages and lock file** — `RestorePackagesWithLockFile=true` enabled in `Directory.Build.props`. Per-project `packages.lock.json` files generated for all projects with dependencies. `verify-quick.ps1` switches to `--locked-mode` when lock files exist. `scripts/test-locked-restore.ps1` performs a deliberate two-run-from-clean-state test: both runs PASS, proving the lock file is honoured without network access. Commits `229e93c` (property + work item), `559212f` (verify + test scripts), `f9c27cf` (lock files), `ab27ee6` (test project lock files). `pwsh ./scripts/verify-quick.ps1` → PASS in ~9 s; `pwsh ./scripts/test-locked-restore.ps1` → PASS.
- **R0.4 — `.editorconfig` and `.gitattributes`** — C# language conventions, analyzer severity, naming rules, var preferences in `.editorconfig`. Line-ending policy (LF for all source, scripting, config, Markdown) and `packages.lock.json` merge strategy in `.gitattributes`. Normalized all 57 tracked text files from CRLF to LF. Added `CA1707` to `NoWarn` with a per-rule comment. `tests/Directory.Build.props` now imports the root props via `GetPathOfFileAbove`, and sets `EnforceCodeStyleInBuild=false` because the test code uses xUnit conventions (snake_case, braceless if) that production-style rules flag. `CoreBoundaryTests.cs` gained `StringComparison.Ordinal` on `Assert.Contains`/`Assert.DoesNotContain` (CA1307). Commits `c66e683` (editorconfig + gitattributes), `50ffabe` (LF normalize), `5f4ac32` (CA1707), `7d64cd7` (test props import + EnforceCodeStyleInBuild off).

## Environment (recorded once, then referenced)

- **Host OS**: Windows 11 25H2 (build 26200), x64, Professional. The registry `ProductName` may report "Windows 10 Pro" — trust the build number, not the string. See `docs/adr/0005-windows-host-build-number.md`.
- **.NET SDK**: 10.0.400
- **.NET runtimes**: 10.0.11, 8.0.30 (Microsoft.NETCore.App, Microsoft.AspNetCore.App, Microsoft.WindowsDesktop.App)
- **Visual Studio**: 18 (Community), installed at `C:\Program Files\Microsoft Visual Studio\18\Community`
- **Office**: Microsoft 365 x64 retail, 16.0.20326.20112, current channel, en-US
- **PowerShell**: 7.4.19
- **Excel-DNA target**: 1.9.0
- **Self-hosted runner policy**: option 1 — a real licensed Microsoft 365 x64 install on this machine; the Office-integration suite is excluded from `verify.ps1` and runs via `scripts/verify-office.ps1` only at phase exit and release.

## Known limitations

- **L2** — `.vs/` (Visual Studio user options) is present locally and is already covered by `.gitignore`.
- **L3** — The optional `03-ROADMAP.md` skill originally referenced a `03` gap; the renumbered `.clinerules/03-ROADMAP.md` now closes that gap and the `docs/03-ROADMAP.md` mirror is committed.

## Open questions for human

- _none yet_
