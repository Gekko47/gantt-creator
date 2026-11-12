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

<!-- SKILL-TOOLS:START -->
- `git_tool` (`status`, `diff`, `log`, `diff --check`) — inspect commits, check for conflict markers, verify Conventional Commit prefixes.
- `github__list_commits` — verify commit history and push status.
- `github__get_pull_request` / `github__get_pull_request_files` / `github__get_pull_request_reviews` — execute the review checklist against actual PR content.
- `github__get_pull_request_status` — verify CI is green before merging (W-12 rule).
- `github__create_pull_request_review` — submit structured reviews.
- `combined-mcp-server__run_commands` with `scripts/install-pre-commit.ps1` — install the pre-commit hook.
- `combined-mcp-server__run_commands` with `scripts/verify-quick.ps1` / `scripts/verify.ps1` — run local gates before commit/PR.
<!-- SKILL-TOOLS:END -->

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
- **Push and pull requests are human-gated.** The agent never pushes
  and never opens a pull request without explicit human instruction.
  The branch lifecycle is: the agent commits (verify-quick green per
  commit, pre-commit hook as the fast safety net); the agent runs
  `verify.ps1` once after its last commit as the branch-final
  self-certification; the branch is pushed for review on instruction;
  the CodeRabbit review stage runs on the pushed branch; the human
  creates the pull request once satisfied. CodeRabbit findings return
  as normal fix commits through the same feature loop.

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

Two gates with different tree-state contracts:

- `verify-quick.ps1` — the every-commit gate. Run it after staging: it
  accepts staged (uncommitted) changes because its skill-tree and
  STATUS checks must validate exactly the content being committed
  (the staged index), not the working tree.
- `verify.ps1` — the branch-final gate. Run it once, after the last
  commit for the branch. Its first step ("working tree hygiene")
  requires a fully clean tree — including no staged changes — and
  fails in under a second on a dirty tree, because the coverage and
  publish artifacts it produces must certify the committed state.

Feature loop (one roadmap row = one commit):

```powershell
git add -A                        # stage everything before the gates
pwsh ./scripts/verify-quick.ps1   # every commit — expect exit 0
git diff --cached --check
git commit                        # the pre-commit hook runs here
# after the LAST commit for the branch:
pwsh ./scripts/verify.ps1         # branch-final — requires a clean tree
```

Before committing, review exactly what will be committed:

```powershell
git diff --cached --check
git diff --cached --stat
git diff --cached
```

## Pre-flight checklist

Before declaring a commit or PR green, run these tools and observe their output:

1. `dotnet_test` on the targeted project — must PASS.
2. `combined-mcp-server__run_commands` with `scripts/verify-quick.ps1` (during editing) or `scripts/verify.ps1` (before PR) — must PASS.
3. `git_tool` with `diff --check` — must show no conflict markers or whitespace drift.
4. `git_tool` with `log --oneline -1` — verify Conventional Commit prefix.
5. `github__get_pull_request_status` (if a PR exists) — must show `success` (W-12 rule).
6. `memra_add` — record one work-item completion fact (3-5 lines: ID, commit, gates, Office status, next item). Not commit text or test counts — see `docs/06-LLM-PROTOCOL.md` "Persisting the ledger".

Pre-commit safety net (optional, recommended):

```powershell
pwsh ./scripts/install-pre-commit.ps1
```

Once installed, every `git commit` runs the fast deterministic gates
(skill-tree drift, STATUS.md accuracy,
markdown-link sanity) before the commit is created. A failure aborts
the commit. The hook deliberately does **not** run `verify-quick.ps1`
(~2-5 minutes build + test; runtime depends on the machine) so a commit
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
    from the version-pinned `scripts/*.ps1` entry points (Pester 4→6+
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
    same entry points verify-quick.ps1 and verify.ps1 call. The CI-only
    package vulnerability scan has its own entry point,
    `scripts/scan-vulnerable-packages.ps1` (W15).
  - Tool pins (Pester minimum major, PSScriptAnalyzer version, actionlint
    SHA-256, actionlint download URL) live in `scripts/tool-versions.psd1`
    and are consumed by `lint-ci.ps1`, `test-scripts.ps1`, and `ci.yml`.
    `ci-parity.Tests.ps1` asserts the two file consumers match the psd1.
  - Every gate must fail loudly when discovery yields zero files or
    items to check (`check-md-links.ps1` now does this for zero scanned
    files), while a clean result with zero findings on a non-trivial,
    discovered input is a PASS. The no-silent-pass rule extends to every
    future gate.

## Gate runtime (W15)

Measured on the reference machine, 2026-11-12:

| Step | Before | After | Cause of the change |
|---|---|---|---|
| status accuracy | 33.8 s | 1.5 s warm, 13.3 s cold | one `git ls-files -z` plus one `git cat-file --batch-check` replace ~440 git processes (239 `rev-parse`, 199 `ls-files`, guide lookups) |
| script lint (Pester) | 141.6 s | 64.4 s | test files run in four parallel runspaces; each check-status fixture is copied from one template repository instead of re-initialised |
| script analyzer | 15.8 s | 16.5 s | **no gain**: the recursive walk is only ~3 s of a ~11 s cold analysis. The analysed set is now explicit and tested instead |
| format | 44.2 s | 44.8 s | **no change**: no equivalent flag set was found (see below) |
| verify-quick total | 295.4 s | 184.4 s | |
| verify total | 431.7 s | not yet re-measured | the vulnerability scan left the local gate for CI |

Rules that keep it fast, each with the evidence that produced it:

- Never reintroduce one process per token in `check-status.ps1`. Batch through
  the index snapshot (`git ls-files -z`) and the hash batch-check
  (`git cat-file --batch-check`); a tag object is still peeled with
  `rev-parse --verify <hash>^{commit}` so `^{commit}` semantics are preserved.
- Keep the Pester throttle at 4 unless a fresh measurement says otherwise.
  Same suite, same machine: 129.8 s at 8 workers, 92.8 s at 4, 94.7 s at 2,
  ~142 s sequential. The files are process-spawn bound, so oversubscribing
  loses. Record the number you measure when changing it. Splitting
  `check-status.Tests.ps1` into two files was tried for the same reason and
  reverted: 65.1 s at four workers before, 63.7 s after. The suite is
  contention-bound now, not bound by one file, so a further split buys nothing
  and costs a duplicated fixture.
- `dotnet format` stays `--verify-no-changes --exclude tests` on the solution.
  `--include src` measured 5.68 s but exited 0 on a deliberately mis-formatted
  file under `src/` (positive control, 2026-11-12) — a false pass, rejected.
  `--no-restore` measured 32.9 s against 31.8 s — no gain. The ~40 s is
  workspace evaluation of all ten projects, including the test projects that
  `--exclude` removes from formatting but not from evaluation.
- The PSSA gate resolves its own file list (`Get-PssaTargetFile`) and fails when
  that list is empty. Do not restore `-Recurse -Exclude '_artifacts'`: the
  explicit list is what makes the analysed set testable. Do not shard it:
  eight `Start-Job` workers measured 21.5 s against 14.3 s serial, because job
  host startup dominates a ~11 s cold analysis.
- The mutation gate must not spend minutes mutating nothing. `Invoke-MutationJudge`
  treats a supplied-but-empty `-InScopeFiles` as an empty changed set and passes
  while stating that reason, and the gate skips Stryker entirely when
  `Get-ChangedCoreFiles` returns nothing. Before that, a docs-or-tooling change
  ran Stryker for 113 s, created 6,811 mutants, scored none, and failed on
  pre-existing Safe Mode compile errors from untouched files.
- `dotnet list package --vulnerable` needs network access, so it runs in CI via
  `scripts/scan-vulnerable-packages.ps1`, not in the offline-capable local
  `verify.ps1`.
- `dotnet-stryker` already selects its own concurrency (`--help`: "By default
  Stryker tries to make the most..."), so the mutation gate does not pin one.

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
