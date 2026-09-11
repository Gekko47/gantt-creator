# Gantt Creator agent contract — revision 3

Applies to every coding agent and every repository change. Reflects the approved architecture: one visible Gantt worksheet plus one `_GanttCreatorConfig` worksheet (`xlSheetVeryHidden`).

Detailed elaboration for each section below lives in the matching on-demand skill under `.cline/skills/`. This file is deliberately short because it is loaded on every turn; the skills are loaded only when the conversation matches them.

## Source-of-truth order

1. Current approved work item in `docs/work-items/`.
2. Product invariants in this file and `docs/02-ARCHITECTURE.md`.
3. Entity contracts in `docs/07-GANTT-ENTITY-GUIDE.md` for any visual, layout, style, label, or export work.
4. Accepted ADRs and `docs/DECISIONS.md`.
5. `docs/03-ROADMAP.md` and `docs/04-TEST-STRATEGY.md`.
6. Existing tested code.
7. External primary documentation.

When sources conflict, stop and report the conflict. Do not silently choose one.

## Non-negotiable product invariants

- Windows desktop Excel add-in: Excel-DNA 1.9, C#, `.NET 10`, `net10.0-windows`, x64 first.
- Exactly one visible user worksheet plus one add-in-managed `xlSheetVeryHidden` configuration worksheet named `_GanttCreatorConfig`. Never create another helper sheet.
- The VeryHidden worksheet stores only versioned catalogues, styles, metrics, label options, defined-name sources, and configuration metadata. Never store activity rows, schedule descriptions/dates, rendered shapes, scene data, calculations, logs, or export staging there.
- The worksheet data table is visible and directly beside the live Gantt.
- The `Type` column uses the single code-owned catalogue in `docs/07-GANTT-ENTITY-GUIDE.md`, materialised to the VeryHidden configuration worksheet for data validation; do not duplicate, extend, or infer type values elsewhere.
- Live output is native Excel cells and shapes. All positioning comes from the scene/layout model.
- Every visible entity uses the geometry, style-token, label, z-order, validation, and renderer-equivalence contract in `docs/07-GANTT-ENTITY-GUIDE.md`. Undefined behaviour is an unknown requiring a decision.
- Normal worksheet, Type, colour, label-position, and selection changes never render. Only the explicit Refresh/render commands rebuild the chart.
- Per-row fill/line/label overrides operate only on one expanded visible entity selection and are stored in that visible table row; they do not directly format the existing generated shape or move schedule data to the configuration sheet.
- The Core project has no Office, Excel-DNA, SkiaSharp, clipboard, filesystem-dialog, or UI dependency.
- Multiple events on one lane, overlapping planned/actual events, milestones, critical intervals, and labelled full-height delineators are first-class domain cases.
- Milestones and delineators read `Start` as their single date and do not use `Finish` for geometry.
- Editable export is an all-shape temporary composition, generated only on request and grouped before copy.
- PowerPoint transfer requests `ppPasteShape` and verifies the returned shape range. Never report success based only on a non-throwing COM call.
- PNG export is generated only on request. Width is user-selected; height preserves the scene aspect ratio; output pixel dimensions and 300-DPI metadata are verified.
- Core operation is offline. No telemetry, web fonts, cloud API, online licence check, or hidden network fallback.

## Required task protocol

Before editing: restate outcome, files/layers likely involved, explicit exclusions, acceptance tests, and remaining unknowns; verify any uncertain API/Office behaviour/package capability against installed metadata or primary documentation before relying on it; ask for direction if the answer changes the product contract, schema, dependency graph, security model, or supported Office versions. Full session-opening format: skill `06-llm-protocol`.

While editing:

- Make the smallest coherent change that proves one behaviour.
- Keep one work item in progress; do not opportunistically refactor neighbours.
- Add or change tests in the same commit as behaviour.
- **Every new `if (bad) { error }` validator ships with a positive test in the same commit** — a `[Fact]` that constructs a bad input and asserts the error path fires, in the same harness that will run it. Validators without positive tests are how this codebase has repeatedly shipped silently-broken failure paths. The pattern: write the validator, write the test that constructs the bad input, both in the same diff. For PowerShell scripts, the equivalent is a self-contained harness in `scripts/` invoked from `verify-quick.ps1`; for C#, a sibling `[Fact]` in the same test project.
- Use the ports/adapters boundaries. Never call Excel or PowerPoint from Core.
- Use deterministic identifiers, ordering, culture, time, and rounding.
- Preserve unrelated user changes and existing public behaviour.
- Never disable analyzers, loosen coverage, delete tests, catch-and-ignore exceptions, or add arbitrary delays to force green output.

After editing: run targeted tests, then `pwsh ./scripts/verify.ps1`; review `git diff --check`, `git status --short`, and the complete diff; update `docs/STATUS.md`; report using the handoff format in skill `06-llm-protocol` (do not claim a command ran unless its output was observed); self-certify against the applicable sections of `docs/08-TEST-CHECKLIST.md` and list the certified sections in the commit message (e.g. `Checklist: A, B, C, G, I`) — do not claim certification for sections that do not apply.

## Stop conditions

Stop and ask the human when:

- acceptance criteria are missing or contradictory;
- a change would violate a product invariant;
- the workbook schema or export fidelity contract needs to change;
- a new production dependency, Office permission, signing certificate, or external service is required;
- a failing test appears to expose existing unrelated behaviour;
- destructive migration or deletion is proposed;
- credentials, signing secrets, customer data, or proprietary fonts are needed;
- the same failure persists after two evidence-led fixes.

## Bounded retry rule

Two failed attempts on the same failure: stop, record the attempts, evidence, current diff, and the smallest question or manual step needed. Never repeat an unchanged command expecting a different result, except once for a documented flaky external Office operation — a flaky retry must be recorded and becomes a defect if it passes only on retry. Full no-loop protocol (failure classification, per-attempt hypothesis discipline, prohibited "fixes"): skill `06-llm-protocol`.

## Evidence and anti-hallucination rules

- Do not invent Excel-DNA callbacks, COM members, RibbonX attributes, enum values, SkiaSharp APIs, package versions, test results, files, or user decisions.
- Prefer compiler/Object Browser/installed package metadata over memory. Prefer vendor documentation over blogs.
- Never say "fully tested" when Office-hosted or manual tests were skipped.

Full evidence-ledger format and verification workflow (fact/inference/proposal/unknown, disposable spikes, citation requirements): skill `06-llm-protocol`.

## C# and architecture rules

- Excel/PowerPoint calls execute on the required STA/main thread. Do not use `Task.Run` around COM.
- Avoid chained COM property calls. Hold each COM proxy in a local variable and release it through one tested ownership helper.

Full rules (nullable/analyzers/immutability conventions, `DateOnly` usage, geometry rounding policy, dependency-injection boundaries, `ScreenUpdating`/`EnableEvents`/`DisplayAlerts` save-restore discipline, error/logging conventions, XML doc policy): skill `02-architecture`.

## Testing rules

- A bug fix starts with a failing regression test, unless the failure exists only inside Office — then add the closest contract test plus an Office reproduction record.
- Golden image updates require human review, a stated reason, and a dedicated commit.
- Do not use `Thread.Sleep` for synchronization. Poll a named observable condition with a deadline and useful timeout diagnostics.

Full policy (coverage thresholds, test layers, Office integration harness, flaky-test policy, traceability map): skill `04-test-strategy`.

## Git and documentation rules

- One concern per commit; do not mix mechanical formatting with behaviour. Commit only a green state, using Conventional Commit prefixes such as `feat:`, `fix:`, `test:`, `refactor:`, `docs:`, `build:`, and `chore:`.
- Do not amend, rebase, force-push, tag, or publish unless explicitly asked.
- Do not commit Office temporary files, build output, test results, exported customer images, or local logs.
- Install and trust the pre-commit hook (`pwsh ./scripts/install-pre-commit.ps1`) — it is a fast safety net (skill-tree drift, STATUS.md accuracy, markdown-link sanity), not a substitute for `verify-quick.ps1`/`verify.ps1`.

Full policy (branch/review policy, CI jobs, PR template, review checklist, comment/dependency policy, release gate): skill `05-git-quality`.

## Tool routing

Use the highest-fidelity tool for the job. Per-skill tool lists live in each `SKILL.md`; this table is the always-on index for the most safety-critical routings:

| When you need to… | Use |
|---|---|
| Get compiler type/symbol info | `vscode-mcp__get_symbol_lsp_info`, `vscode-mcp__get_diagnostics` |
| Find an API's real signature | `context7__query-docs`, `microsoft-learn__microsoft_docs_search`, `ilspy_decompile` |
| Run tests / check coverage | `dotnet_test`, `dotnet_build` |
| Verify a NuGet package claim | `dotnet_packages` |
| Verify green CI on a branch | `github__get_pull_request_status` |
| Persist a decision or evidence across turns | `memra_add_decision`, `memra_add` |
| Recall prior decisions at session start | `memra_bootstrap` |
| Structure multi-step reasoning | `sequential-thinking__sequentialthinking` |
| Check an architecture invariant | `roslyn_analyze`, `dotnet_test` on `*.Architecture.Tests` |
| Find a pattern violation across the codebase | `ast_grep_search`, `semgrep_scan` |

## Completion language

Use short, factual handoffs. Avoid praise, speculation, marketing language, and exhaustive file-by-file narration. If a gate was not run, state `Not run` and why.
