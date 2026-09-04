# Gantt Creator agent contract — revision 3

> Created 2 September 2026 under a new filename. This revision incorporates the approved architecture of one visible Gantt worksheet plus one `_GanttCreatorConfig` worksheet with `xlSheetVeryHidden` visibility. When installed in the repository, rename this file to `AGENTS.md`.

This file applies to every coding agent and every repository change.

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

Before editing:

1. Read the active work item and only the linked architecture/test sections.
2. Inspect the relevant implementation and tests.
3. Restate: intended outcome, files likely to change, exclusions, acceptance tests, and uncertainties.
4. If an API, Office behaviour, or package capability is uncertain, verify it in installed metadata, a minimal spike, or primary documentation. Mark an unverified statement as a hypothesis.
5. Ask for direction if the answer changes the product contract, public data schema, dependency graph, security model, or supported Office versions.

While editing:

- Make the smallest coherent change that proves one behaviour.
- Keep one work item in progress; do not opportunistically refactor neighbours.
- Add or change tests in the same commit as behaviour.
- Use the ports/adapters boundaries. Never call Excel or PowerPoint from Core.
- Use deterministic identifiers, ordering, culture, time, and rounding.
- Preserve unrelated user changes and existing public behaviour.
- Never disable analyzers, loosen coverage, delete tests, catch-and-ignore exceptions, or add arbitrary delays to force green output.

After editing:

1. Run targeted tests.
2. Run `pwsh ./scripts/verify.ps1`.
3. Review `git diff --check`, `git status --short`, and the complete diff.
4. Update task evidence and `docs/STATUS.md`.
5. Report exactly: changed behaviour, important files, commands run and results, residual risks, and the next roadmap item. Do not claim a command ran unless its output was observed.

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

Do not loop.

1. First failure: capture the exact command, error, environment, and likely layer. Form one falsifiable hypothesis and run one discriminating check.
2. Second failure: re-read the relevant code/documentation, change the hypothesis, and try one materially different fix.
3. Third occurrence: stop. Record attempts, evidence, current diff, and the smallest question or manual step needed.

Never repeat an unchanged command expecting a different result, except once for a documented flaky external Office operation. A flaky retry must be recorded and becomes a defect if it passes only on retry.

## Evidence and anti-hallucination rules

- Do not invent Excel-DNA callbacks, COM members, RibbonX attributes, enum values, SkiaSharp APIs, package versions, test results, files, or user decisions.
- Prefer compiler/Object Browser/installed package metadata over memory. Prefer vendor documentation over blogs.
- Cite the API or test proving any non-obvious interoperability claim in the work item.
- Use a minimal disposable spike when documentation cannot answer a compatibility question. Do not merge the spike as product code.
- Distinguish `fact`, `inference`, `proposal`, and `unknown` in investigation notes.
- Never say “fully tested” when Office-hosted or manual tests were skipped.

## C# and architecture rules

- Nullable reference types, implicit usings, analyzers, and warnings-as-errors remain enabled.
- Prefer immutable records/value objects in Core. Use `DateOnly` for date-only schedule data.
- Define geometry in points with central rounding/tolerance policies. Do not scatter pixel conversions or magic offsets.
- Inject clock, file dialog, clipboard, Office application, and logging boundaries.
- Avoid `dynamic` unless an isolated, documented late-binding compatibility adapter requires it.
- Avoid chained COM property calls. Hold each COM proxy in a local variable and release it through one tested ownership helper.
- Excel/PowerPoint calls execute on the required STA/main thread. Do not use `Task.Run` around COM.
- Save and restore `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, calculation mode, status bar, and selection only when changed, using `try/finally`.
- User errors are concise and actionable. Technical details go to a local rolling log with no workbook content unless explicitly opted in.
- Public APIs need XML documentation when the contract is not obvious. Internal comments explain why, invariants, or Office quirks—not line-by-line mechanics.

## Testing rules

- Test observable behaviour, not private method implementation.
- A bug fix starts with a failing regression test unless the failure exists only inside Office; then add the closest contract test plus an Office reproduction record.
- Core and scene/layout tests must be deterministic and parallel-safe.
- Golden image updates require human review, a stated reason, and a dedicated commit.
- Office integration tests are tagged `OfficeIntegration`, serialize access, start from clean fixture files, and clean up processes/artifacts in `finally`.
- Do not use `Thread.Sleep` for synchronization. Poll a named observable condition with a deadline and useful timeout diagnostics.

## Git and documentation rules

- One concern per commit; do not mix mechanical formatting with behaviour.
- Commit only a green state. Use Conventional Commit prefixes such as `feat:`, `fix:`, `test:`, `refactor:`, `docs:`, `build:`, and `chore:`.
- Do not amend, rebase, force-push, tag, or publish unless explicitly asked.
- Do not commit Office temporary files, build output, test results, exported customer images, or local logs.
- Update comments only when they add current, non-obvious value. Remove stale comments in the touched area.
- Keep the work item and status concise; they are control records, not diaries.

## Completion language

Use short, factual handoffs. Avoid praise, speculation, marketing language, and exhaustive file-by-file narration. If a gate was not run, state `Not run` and why.
