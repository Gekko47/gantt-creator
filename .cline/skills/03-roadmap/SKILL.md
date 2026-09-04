---
name: 03-ROADMAP
description: Use this skill when the conversation is about which phase, which work item (R0.x .. R10.8), which automated gate, or which Office gate applies. Also use it for the cross-phase compatibility matrix, the scope-change protocol, and when the user asks for the next safe action.
---

# Implementation roadmap — revision 3
> Created 2 September 2026 under a new filename. This revision contains 96 commit-sized work items and adds creation, named-range catalogues, migration, safe repair, and verification for the approved VeryHidden configuration worksheet. When installed, use the path `docs/03-ROADMAP.md`.
## How to use this roadmap
- Complete items in order unless an approved ADR records why order changed.
- One row is normally one commit. Split a row if the diff becomes difficult to review; do not combine rows merely because they are related.
- Create a work-item file before implementation. Record the exact acceptance tests and evidence.
- Every work-item acceptance criterion that names a test count, a command, or an artifact must link to a specific test or script step. Drift between the doc and the code is a defect (see `docs/08-TEST-CHECKLIST.md` section I).
- Every visual/layout/style/export work item names the affected sections of `docs/07-GANTT-ENTITY-GUIDE.md` and tests the defined cross-renderer contract.
- Every code commit must pass `scripts/verify-quick.ps1`; every phase exit and pull request must pass `scripts/verify.ps1`.
- A phase exits only after its stated demonstration. A screenshot is supporting evidence, not a substitute for automated assertions.
- Use synthetic construction data. Do not add customer schedule data to tests or examples.
The **Automated gate** is run from the terminal and CI. A **Visual Studio / Office gate** marked `Required` must be demonstrated against desktop Excel or PowerPoint on Windows, normally launched or debugged through Visual Studio. `None` means the commit does not require Office; it does not waive the automated gate. A work item cannot be marked done when a required Office gate was not run.
## Phase 0 — repository and quality foundation
Goal: a clean solution that fails fast on warnings, formatting, dependency drift, and broken tests.
| ID | Reviewable commit outcome | Automated gate | Visual Studio / Office gate |
| --- | --- | --- | --- |
| R0.1 | Add governance kit, licence, security policy, and contribution entry point | Validate Markdown links, rule discovery, and skill schemas | None |
| R0.2 | Add `global.json`, solution, project folders, and allowed project references | Release build and architecture reference test | Required: open the solution and confirm every project and x64 configuration loads |
| R0.3 | Add central package management, locked restore, and approved initial packages | Locked restore succeeds twice from clean package state | None |
| R0.4 | Add `.editorconfig`, `Directory.Build.props`, analyzers, nullable, and warnings-as-errors | Format check plus a temporary warning that must fail before removal | None |
| R0.5 | Add xUnit test projects and one non-trivial sample test per test project | `dotnet test` discovers and passes every test project | Required: Test Explorer discovers the same test projects |
| R0.6 | Add quick/full verification scripts and coverage settings | Both scripts pass from a clean clone and enforce configured thresholds | None |
| R0.7 | Add Windows CI for restore, format, build, unit tests, and artifacts | Branch workflow passes and a deliberate failure produces a useful annotation | None |
| R0.8 | Add versioning and local rolling-log abstractions with privacy-safe defaults | Unit tests for version string, rotation, and redaction | None |
Exit demonstration: a new developer clones the repository, runs one documented command, and gets a clean Release build and test report without opening Office.

---

## Where to read more

- Canonical source: `docs/03-ROADMAP.md`
- Always-on rule: `.clinerules/03-ROADMAP.md`
- Full reference: `./references.md` in this directory
