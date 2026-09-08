# Git, quality gates, and review


<!-- SKILL-SUMMARY:START -->
Branch/review policy, commit design, local gates, CI parity, and the
pull-request/review checklist.

Do not get wrong:
- A phase may not exit on local-only evidence: the branch must be
  pushed and GitHub CI must have been observed green (the W-12 rule —
  added after the PSSA blind-gate defect survived three local-only
  "all green" QA rounds).
- Conventional Commit prefixes only (`feat:`, `fix:`, `test:`,
  `refactor:`, `docs:`, `build:`, `chore:`); one concern per commit.
- Never rewrite shared history (amend/rebase/force-push) without
  explicit instruction.
- Comments only when they add current, non-obvious value; remove stale
  ones in the touched area. Work items and STATUS are control records,
  not diaries.
<!-- SKILL-SUMMARY:END -->

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

## Control-record discipline

Keep the work item and status concise; they are control records, not diaries.

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
  - Every gate must fail loudly when discovery yields zero files or
    items to check (`check-md-links.ps1` now does this for zero scanned
    files), while a clean result with zero findings on a non-trivial,
    discovered input is a PASS. The no-silent-pass rule extends to every
    future gate.

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
- Logs exclude user schedule content by default.

### Readability

- Names make the normal path clear.
- Comments explain rationale or an interoperability constraint and are still accurate.
- XML documentation describes real public contracts, not implementation details.
- The diff contains no commented-out code, stale TODO, or generated noise.

## Comment policy

Good comment:

```csharp
// Excel reports shape coordinates as Single values, so compare after one
// point-space rounding step rather than repeatedly converting to pixels.
```

Bad comments:

```csharp
// Loop through the events.
// Set width.
// This method renders the Gantt chart.
```

Use issue-backed TODOs only when deferral is intentional: `// TODO(#123): ...`. Do not leave speculative notes or describe code that could be named more clearly.

## Dependency policy

A production package addition requires:

- capability not reasonably supplied by .NET/Office/existing code;
- official project/source link and current maintenance evidence;
- compatible commercial licence and notice requirements;
- offline deployment/native binary analysis;
- pinned version and lock-file update;
- vulnerability scan;
- an ADR when it materially shapes architecture or output fidelity.

An agent may propose a dependency but may not add it without task scope or human approval.

## Release gate

The release candidate requires:

- clean reproducible build from a tagged commit;
- all unit/contract/architecture/golden tests passing;
- supported Office integration matrix passing;
- performance and repeated-operation tests within accepted budgets;
- offline acceptance run;
- signed binaries and verified signature;
- SBOM, third-party notices, checksums, and version notes;
- installer/update/rollback/uninstall rehearsal on a clean VM;
- known limitations reviewed and accepted by a human.

Do not label a build production-ready because it works on the developer's Excel installation.
