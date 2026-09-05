---
name: 05-GIT-QUALITY
description: Use this skill when the conversation is about branch and review policy, commit design and Conventional Commit format, the local gates (verify-quick.ps1 / verify.ps1), CI jobs, the pull-request template, the review checklist, the comment policy, the dependency policy, or the release gate.
---

# Git, quality gates, and review
## Branch and review policy
- Protect `main`: pull requests, passing required checks, and one human approval.
- Name branches `type/roadmap-id-short-description`, for example `feat/r3-09-stacked-events`.
- Rebase/merge policy is a team choice; do not let an agent rewrite shared history.
- No direct production release from an unreviewed local working tree.
## Commit design
Each commit:
- implements one roadmap outcome or one isolated preparatory refactor;
- keeps the solution buildable and tests green;
- includes tests with changed behaviour;
- avoids unrelated formatting or rename noise;
- updates only the necessary control documents;
- is understandable without reading later commits.
Suggested soft limit: 50-250 changed production lines plus tests. Split by behaviour, not by arbitrary file count. A pure rename, generated interop file, or approved golden-image commit can exceed the limit but must be isolated.
Commit format:
```text
feat(scene): render stacked events on one lane
Map StackIndex to deterministic sub-lane offsets while preserving lane height.
Tests cover shuffled input, duplicate stack values, and clipping.
```
Do not use messages such as `updates`, `fix stuff`, or an agent transcript.
## Local gates
During editing:
```powershell
pwsh ./scripts/verify-quick.ps1
```
Before every commit/PR:
```powershell
pwsh ./scripts/verify.ps1
git diff --check
git status --short
git diff --stat
git diff
```
Quick verification runs format check, Release build, and non-Office tests without coverage packaging. Full verification runs locked restore, format/analyzers, Release build, all non-Office tests with configured coverage thresholds, and repository hygiene checks.
Do not commit when a gate is red. Do not bypass the script by running only the test that passes.
## CI jobs
Required pull-request jobs:
1. repository hygiene and secret scan;
2. locked NuGet restore;
3. `dotnet format --verify-no-changes`;
4. Release build with warnings-as-errors;
5. Core, Raster, Office contract, and AddIn tests;
6. per-project coverage thresholds;
7. architecture/dependency tests;
8. golden image comparison on a pinned Windows runner;
9. package vulnerability and licence report.
Office integration runs on a controlled self-hosted Windows runner at phase exits, nightly if stable, and always before a release candidate. It is not replaced by unit coverage.
Cache only NuGet packages keyed by lock files. Never cache build outputs in a way that can hide a clean-build failure.
## Pull-request description
Keep it brief and evidence-led:
```markdown
## Outcome
<observable behaviour>
## Scope
- <important change>
- <important change>
## Proof
- `<command>` — PASS
- Office/visual check — PASS / Not run: <reason>
## Risk and rollback
<main risk and simple rollback>
## Excluded
<nearby work intentionally not done>
```
## Review checklist
### Behaviour
- Work item acceptance criteria are met and no extra product behaviour appeared.
- Error/empty/boundary cases are visible in tests.
- Existing one-sheet and offline constraints remain true.
### Architecture
- Core has no infrastructure reference.
- Scene layout is not duplicated in an Office or raster renderer.
- New COM calls are isolated and have clear ownership/state restoration.
- New configuration has versioning and a migration/default policy.
### Quality
- Assertions would fail for a plausible defect.
- No suppressed warning, reduced threshold, hidden retry, arbitrary sleep, or catch-and-ignore.
- No unnecessary dependency or public API.

---

## Where to read more

- Canonical source: `docs/05-GIT-QUALITY.md`
- Always-on rule: `.clinerules/05-GIT-QUALITY.md`
- Full reference: `./references.md` in this directory
