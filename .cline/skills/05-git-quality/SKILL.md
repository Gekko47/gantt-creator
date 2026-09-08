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
- **A phase may not exit on local-only evidence.** The branch must be
  pushed and the GitHub CI gate must have been observed green for the
  work item. Local PASSes alone are insufficient — the W-12 amendment
  was added after the PSSA blind-gate defect survived three QA rounds
  because CI had never been run on the branch.
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
Pre-commit safety net (optional, recommended):
```powershell
pwsh ./scripts/install-pre-commit.ps1
```
Once installed, every `git commit` runs the fast deterministic gates
(skill-tree drift, SKILL.md canonical phrases, STATUS.md accuracy,
markdown-link sanity) before the commit is created. A failure aborts
the commit. The hook deliberately does **not** run `verify-quick.ps1`
(~2 minutes build + test; runtime depends on the machine) so a commit
is not slowed down; the developer still runs `verify-quick.ps1` during
editing and `verify.ps1` before a PR.
See `AGENTS.md` and `scripts/pre-commit.ps1` for the full contract.
Quick verification runs format check, Release build, and non-Office tests without coverage packaging. Full verification runs locked restore, format/analyzers, Release build, all non-Office tests with configured coverage thresholds, and repository hygiene checks.
Do not commit when a gate is red. Do not bypass the script by running only the test that passes.
## CI parity (W8)
The CI workflow and the local verify scripts must agree on every step. The
defect class "local says PASS, CI says FAIL because the two views diverged"
includes:
  - inline `dotnet test`/`dotnet build`/`dotnet publish`/`Invoke-Pester`
    calls in `ci.yml` that drift
    from the version-pinned `scripts/*.ps1` entry points (Pester 4→5
    removed the `-Script` parameter; the first CI push of `stage-inspect`
    would have failed on it);
  - a tool pin declared in two places (e.g. actionlint SHA-256 in
    both `ci.yml` and `lint-ci.ps1`) that is patched in one but not the
    other;
  - a script-side gate that silently passes locally but fails on the
    CI runner (e.g. the Phase-C PSScriptAnalyzer `-EnableExit` defect).
Rules:
  - `.github/workflows/ci.yml` delegates every step to a
    `scripts/*.ps1` entry point or to a vetted native MSBuild command
    (`dotnet format`, `dotnet restore --locked-mode`). No inline Pester
    or `dotnet test`/`dotnet build`/`dotnet publish` invocations: those
    commands live in `scripts/test-non-office.ps1`,
    `scripts/build-release.ps1`, and `scripts/publish-addin.ps1`, the
    same entry points verify-quick.ps1 and verify.ps1 call.
  - Tool pins (Pester minimum major, PSScriptAnalyzer version, actionlint
    SHA-256, actionlint download URL) live in `scripts/tool-versions.psd1`
    and are consumed by `lint-ci.ps1`, `test-scripts.ps1`, and `ci.yml`.
    `ci-parity.Tests.ps1` asserts the two file consumers match the psd1.

---

## Where to read more

- Canonical source: `docs/05-GIT-QUALITY.md`
- Full reference: `./references.md` in this directory
