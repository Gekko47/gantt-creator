# LLM operating protocol


<!-- SKILL-SUMMARY:START -->
The LLM operating protocol for single-work-item tasks (session
opening, evidence ledger, anti-drift/anti-hallucination controls, the
two-attempt no-loop rule, coding behaviour, prompt pattern, handoff
format) AND a separate methodology for a full-repository code review
(six phases: read the source, probe empirically, classify defects
into the four repo-specific classes, plan, implement/verify/commit,
self-apply).

Do not get wrong:
- Two failed attempts on the same failure = stop and report; never a
  third unchanged retry — except exactly one recorded retry for a
  documented flaky external Office operation (AGENTS.md).
- An inference cannot become a fact by being repeated — use the
  `fact`/`inference`/`proposal`/`unknown` ledger honestly.
- A full-repo review must look for gate-integrity, build-pipeline
  drift, docs/code drift, and local/CI divergence even when the
  per-item checklists already pass — those are cross-cutting defect
  classes the per-item gates do not catch by construction.
<!-- SKILL-SUMMARY:END -->

<!-- SKILL-TOOLS:START -->
- `memra_add` / `memra_add_decision` — persist evidence-ledger rows and irreversible decisions across turns.
- `memra_bootstrap` — recall prior decisions and patterns at session start.
- `memra_add_pattern` — store reusable methodologies (e.g. the full-repo review phases).
- `memra_search` / `memra_recall` — retrieve prior evidence by semantic similarity.
- `memra_supersede` — update an evidence-ledger row when its status changes, preserving the history chain.
- `sequential-thinking__sequentialthinking` — structure multi-step reasoning (review methodology, failure classification).
- `vscode-mcp__get_symbol_lsp_info` / `vscode-mcp__get_references` — get compiler-grade type/symbol info instead of guessing APIs.
- `vscode-mcp__get_diagnostics` — fast quality-check for modified files (errors, warnings, hints); much faster than `tsc --noEmit` or `eslint .`.
- `ilspy_decompile` — inspect installed assembly metadata when source is unavailable.
- `context7__query-docs` / `microsoft-learn__microsoft_docs_search` / `microsoft-learn__microsoft_docs_fetch` — look up unfamiliar .NET / Excel-DNA / SkiaSharp APIs; preferred over memory.
- `dotnet_build` / `dotnet_test` — compile and run tests; "Not run" must be stated if skipped.
- `pwsh_run` — execute PowerShell scripts and commands (verify gates, lint, sync) with observed output.
- `semgrep_scan` / `ast_grep_search` — find pattern violations during review.
<!-- SKILL-TOOLS:END -->

## Purpose

This protocol makes an LLM useful as a bounded engineering assistant. It does not delegate product ownership, evidence, release authority, or irreversible actions.

## Required session opening

Give the agent one work-item ID. The agent must respond with:

1. one-sentence outcome;
2. files/layers likely involved;
3. explicit exclusions;
4. acceptance tests it will use;
5. facts already verified;
6. remaining unknowns;
7. first safe action.

If this answer is wider than the work item, correct it before allowing edits.

## Evidence ledger

For interoperability investigations, maintain this compact table in the work item:

| Statement | Status | Evidence |
| --- | --- | --- |
| `Shapes.PasteSpecial` can request `ppPasteShape` | fact | Microsoft API documentation |
| copied group remains editable after staging cleanup | unknown until tested | compatibility spike R6.7 |
| exact live screen pixels are invariant across zoom/scaling | false | Excel uses point geometry and host rasterisation |

Allowed statuses are `fact`, `inference`, `proposal`, and `unknown`. An inference cannot become a fact because the agent repeats it.

### Persisting the ledger

The evidence ledger is a cross-turn record. Persist every row so it survives session restarts:

- Every `fact` row → `memra_add` with `type: "fact"`, `importance: 8`, and the evidence URL or command in metadata.
- Every irreversible decision → `memra_add_decision` with the context that made it irreversible.
- Every reusable methodology (e.g. the full-repo review phases) → `memra_add_pattern`.
- Session start → `memra_bootstrap` to recall prior decisions before restating unknowns.
- When a statement's status changes → `memra_supersede` to keep the history chain intact.

## Anti-drift controls

- One active work item; one acceptance boundary.
- Fixed decisions live in architecture/ADRs, not chat memory.
- The agent reads only relevant files, then names them in its plan.
- Any new requirement is placed in a later work item unless necessary for the current acceptance test.
- Any public schema, dependency, supported-platform, or export-contract change stops for human approval.
- Refactoring is allowed only when the current slice needs it; preparatory refactors are separate green commits.
- At the end, compare the diff against the initial file/scope list and explain every deviation.

## Anti-hallucination controls

For every unfamiliar API or version-sensitive claim, use this tool order (highest fidelity first):

1. `vscode-mcp__get_symbol_lsp_info` / `vscode-mcp__get_references` — compiler-grade type and symbol info; the primary inspection path.
2. `ilspy_decompile` — inspect installed assembly metadata when source is unavailable.
3. `context7__query-docs` / `microsoft-learn__microsoft_docs_search` / `microsoft-learn__microsoft_docs_fetch` — official .NET, Excel-DNA, SkiaSharp, or Microsoft documentation; preferred over memory and blogs.
4. `dotnet_build` — compile a minimal usage to verify the API exists and behaves as claimed.
5. If host behaviour remains uncertain, write a disposable spike with one measurable assertion (`dotnet_new` + `dotnet_test`).
6. Keep the statement `unknown` until evidence exists.

The agent must not fabricate command output. If it cannot run a test, it says `Not run` and provides the exact command the human should run.

## No-loop protocol

Classify the failure first: compile, deterministic test, Office host, environment, permission, file lock, dependency, or requirement ambiguity.

Attempt 1:

- preserve exact error/HRESULT/stack trace;
- form one falsifiable hypothesis;
- run one discriminating check;
- make the smallest corresponding fix;
- rerun the narrow failed command.

Attempt 2:

- re-read the relevant boundary and primary documentation;
- form a materially different hypothesis;
- inspect environment/process/loaded-module evidence where relevant;
- make one different fix and rerun.

If the same failure occurs again, stop. The sole exception is one
recorded retry of a documented flaky external Office operation, per
the bounded retry rule in AGENTS.md; the retry must be recorded and
becomes a defect if it passes only on retry. Report:

- exact command and environment;
- shortest useful error excerpt;
- two hypotheses, tests, and outcomes;
- current uncommitted diff;
- likely layer and remaining unknown;
- one question or manual observation needed.

Prohibited “fixes” include repeated unchanged commands, larger sleeps, broad exception catches, test deletion, threshold reduction, warning suppression, package churn, random API substitution, and restarting everything without recording what the restart tests.

## Context discipline

- Start a fresh agent session for each roadmap item or after a large context shift.
- Keep `docs/STATUS.md` as the handoff; do not paste whole chat transcripts into the repository.
- Ask the agent to summarise current facts before context compaction.
- Do not feed customer workbooks or proprietary schedule descriptions to a hosted model.
- Use synthetic/minimised reproductions; redact file paths, names, dates, and labels from logs/prompts where needed.

## End-to-end code review methodology

A full code review across the whole repository is a different exercise
from a single-work-item implementation: the reviewer must find the
defect classes that survive individual work-item gates, not the
defects a single work item could have shipped. This section is the
methodology, with worked references to the W-13 / Phase A/B/C run on
this repo (2026-09-07) as the example.

The six phases, in order:

### 1. Read the source (line by line)

Open the files. Read them. Do not rely on docs, on the README, on
prior conversations, on what you remember of the codebase. The point
of this phase is to know what the code *is*, not what someone said it
is. The docs and the code are two separate things; they may disagree,
and the defect you are looking for is often in the disagreement.

For this repo the source-of-truth list at R0.8 is:
- `src/GanttCreator.Core/` (3 files + Logging namespace at the time
  of the W-13 run)
- `src/GanttCreator.Raster/` (1 file)
- `src/GanttCreator.Office/` (empty)
- `src/GanttCreator.AddIn/` (empty)
- `tests/` (6 projects)
- `scripts/` (20 PowerShell files)
- `.github/workflows/ci.yml` (1 file)
- `Directory.Build.props` + `Directory.Packages.props` + `global.json`

Read every line that is not a generated file. The session is bounded
by what the file count allows; the W-13 run read every one of the
above in one pass.

### 2. Probe empirically

For every defect hypothesis, run a discriminating check *now*, in this
session. Do not assume memory, do not assume the documentation, do
not assume a PSSA warning is benign because someone disabled it once.
The probes that produced the W-13 findings:

- Direct invocation of the failing step, outside the verify script:
  `pwsh -Command "Invoke-ScriptAnalyzer ... -Settings ... -EnableExit"`
  and observed exit code 11 with 11 warnings. **The local gate had
  been reporting PASS for the entire R0.8 round; the probe exposed
  the blind-gate defect immediately.** (Count: this doc earlier said
  10; `docs/STATUS.md` and `docs/KNOWN-LIMITATIONS.md` L10 both record
  the verified 11, which is the number used here — reported rather
  than silently reconciled.)
- `Get-Module -ListAvailable PSScriptAnalyzer` to confirm the local
  module version matched CI's pin (1.25.0). The probe was small and
  the answer was decisive.
- `Get-Command Invoke-Pester` to inspect the parameter surface;
  output showed the `Script` parameter does not exist in Pester 6+.
  **This was the proof that the CI workflow's inline
  `Invoke-Pester -Script` call could not work.**
- `(Get-Content ... -Raw) -split "`n" | Select-String` to scan the CI
  workflow and the lint-ci script for the actionlint hash. A SHA-256
  pin declared in two files that disagrees is a drift defect.

Every probe is a single command with an observed output. If a probe
returns a result you cannot interpret, that is a finding, not a
dead-end.

### 3. Identify defect classes

The four classes the W-13 run caught, in increasing severity:

- **A. Gate integrity.** A gate that reports PASS while its target
  failure exists. The PSScriptAnalyzer step was blind; every other
  script gate is audited for the same defect (Phase C W8/W9/W13 work
  hardened them all). **No script gate in this repo now relies on
  `-EnableExit`, function-level `throw`, or any signal that does not
  propagate through a pipeline.** Enforced by
  `scripts/pssa-gate.Tests.ps1` (the positive-control test that would
  catch a regression).
- **B. Build pipeline drift.** A test that reads a `bin/` or
  `publish/` artifact without naming the verify-script step that
  produces it. On a fresh clone, the test could pass on a stale
  artifact and the reviewer would not know. Enforced by
  `tests/.../Architecture.Tests/ArtifactSourceMarkerTests.cs`.
- **C. Docs/code drift.** A verify script's `.DESCRIPTION` block
  enumerated the steps it runs, but nothing checked that the
  enumeration matched the actual `Invoke-Step` calls. The reviewer
  could be told "this script runs step 1, step 2, step 3" while the
  script actually ran step 1, step 2, step 4. Enforced by
  `scripts/step-parity.Tests.ps1`.
- **D. Local/CI view divergence.** A test, tool pin, or behaviour
  that was declared in two places (ci.yml + a local script) and
  patched in one but not the other. On this repo the actionlint
  SHA-256 and the PSScriptAnalyzer version were both such pins.
  Enforced by `scripts/tool-versions.psd1` as the single source of
  truth and `scripts/ci-parity.Tests.ps1` for the three-way equality
  assertion.

A full code review **must look for these four classes even when
checklist A/B/C/D passes**. The checklists catch the per-item
defects; the methodology catches the cross-cutting ones.

### 4. Plan

Apply the AGENTS.md task protocol to each defect class:

- One work item per defect class. Mixing them creates a commit that
  is impossible to bisect and hard to review.
- For each item: restated outcome, files likely to change, exclusions,
  acceptance tests, and uncertainties. If the answer is wider than
  the work item, correct it before editing.
- Where the fix touches a public contract, schema, dependency, or
  supported-platform: stop and ask the human. The W-13 / Phase C
  amendments did not touch any of those, so the human approved the
  default policies.
- Use the two-attempt no-loop rule. Every defect class above was
  caught on the first probe; some required two attempts to fix
  (the rolling-log rotation matcher, the check-md-links empty-scan
  test). Three failures with no new evidence = stop and report.

### 5. Implement, verify, commit

- Every behavior change that adds a validator (`if (bad) { error }`)
  ships with a positive test in the same commit: a `[Fact]` (or the
  script-test equivalent) that constructs the bad input and asserts the
  error path fires. `scripts/pssa-gate.Tests.ps1`,
  `scripts/step-parity.Tests.ps1`, and
  `tests/.../Architecture.Tests/ArtifactSourceMarkerTests.cs` are the
  three worked examples for validator behaviour; they follow the same
  shape. Documentation-only changes do not need a validator positive
  test; they carry the appropriate documentation, link, or drift test
  instead (for example the skill-tree, STATUS, and markdown-link gates
  in verify-quick.ps1).
- Run the narrow test after each meaningful edit. The Phase C
  artefacts discovered two regressions this way (the W-9
  `check-md-links` empty-scan test failed first run; the W-11 ci.yml
  PSScriptAnalyzer literal-vs-variable was caught by the ci-parity
  tripwire).
- Run `pwsh ./scripts/verify-quick.ps1` once the slice is coherent,
  not after every edit. The 12 steps take ~2-5 minutes; running them
  per-edit is wasteful, and skipping them before commit has caught
  regressions in the W-13 run.
- Commit one concern per commit. Multiple concerns in one commit
  produce a confusing diff and a misleading message; the W-13 run
  produced seven commits for two defect classes (Phase A + Phase C)
  and the human could review each independently.
- The branch must be pushed and the GitHub CI gate observed green
  before declaring the review done. Local PASSes are not enough
  (this is the W-12 rule, motivated by the fact that the PSSA
  blind-gate survived three rounds of local-only "all green"
  reports). See the "Branch and review policy" section of
  `docs/05-GIT-QUALITY.md`.

### 6. Self-application

This methodology was used to produce the W-13 / Phase A/B/C run on
2026-09-07. The closing evidence is two green CI runs of `stage-inspect`
on a branch that included every defect-class fix. A third CI run
that observes the same green after a fresh clone is the strongest
signal that the methodology produced durable, reproducible gates.
A third green run remains the strongest signal, but failing to observe
it is not automatically a methodology failure. Classify the failure from
observed evidence first: runner, permission, network, or tooling
failures are not repository-caused and do not by themselves impeach the
probes or the positive-control tests. Only when the evidence points to
the repository (a probe or positive-control test failing on a clean,
supported environment) should the probes and positive-control tests be
reviewed. If the evidence is inconclusive, record the run as `unknown`
with the exact command, environment, and output rather than defaulting
to either conclusion.

## Coding behaviour

The agent should:

- inspect before editing;
- use existing patterns when tested and consistent with current architecture;
- prefer types and small functions to comments;
- add tests before or with the behaviour;
- run the narrow test after each meaningful edit;
- run full verification once the slice is coherent;
- leave the repository cleaner only within the touched scope.

The agent should not:

- design the whole system again in each task;
- create abstraction layers without two concrete consumers or a clear port boundary;
- generate speculative compatibility fallbacks;
- rewrite working code for style alone;
- introduce “temporary” hidden worksheets;
- claim pixel perfection for live Excel screen display;
- terminate Office processes it did not start.

## Prompt pattern

Use:

```text
Implement work item <ID> from docs/work-items/<file>.md.
Follow AGENTS.md and use the relevant project skill.
Before editing, restate outcome, exclusions, tests, facts, and unknowns.
Do not change architecture or dependencies.
Use the two-attempt no-loop rule.
Run targeted tests and scripts/verify.ps1, then report observed evidence only.
Do not commit until I review the diff.
```

Avoid broad prompts such as “build the Gantt app”, “finish the phase”, or “fix all errors”.

## Human review rhythm

The human reviews:

- the initial boundary before edits;
- the diff after each work item;
- every package addition and ADR;
- every golden-image change;
- every Office-hosted visual/clipboard/PowerPoint gate;
- each commit message;
- phase-exit and release evidence.

If review repeatedly finds the same error, update the relevant always-on rule or focused skill. Do not inflate every prompt with the lesson.

## Agent handoff format

```text
Outcome: <what now works>
Changed: <2-5 important files/areas>
Proof: <commands and observed results>
Office/manual: <PASS, FAIL, or Not run: reason>
Risk: <one sentence or None known>
Deferred: <explicitly excluded work>
Next: <roadmap ID>
```

This is intentionally short. The code, tests, work item, and Git diff are the durable record.

## Pre-flight checklist

Before declaring a work item done, run these tools and observe their output:

1. `dotnet_test` on the targeted project — must PASS.
2. `pwsh_run` with `scripts/verify-quick.ps1` — must PASS (the every-commit
   gate; run it with the work staged — it accepts staged changes).
3. `pwsh_run` with `scripts/verify.ps1` — the branch-final gate, run once
   after the LAST commit for the branch; requires a clean tree (its first
   step fails in under a second otherwise). The agent never pushes and
   never opens a pull request without explicit human instruction; the
   human creates the PR after the CodeRabbit review stage (see
   `docs/05-GIT-QUALITY.md`, "Branch and review policy").
4. `vscode-mcp__get_diagnostics` on modified files — must show zero errors.
5. `memra_add` — persist the evidence ledger rows for this session.
6. `memra_add_decision` (if an irreversible decision was made) — persist it with context.
7. `sequential-thinking__sequentialthinking` — use for any multi-step reasoning (failure classification, review methodology) before concluding.
