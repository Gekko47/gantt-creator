# Copilot custom instructions

These instructions apply to GitHub Copilot in Visual Studio when Agent Mode is enabled. They mirror the always-on rules in `.clinerules/00..07-*.md` and `AGENTS.md`. When they conflict, the kit wins; the human reviews the diff.

## Read these first

- `AGENTS.md` — the always-on operating contract for every agent and every commit.
- `docs/01-ENVIRONMENT.md` — the supported baseline and common failures.
- `docs/02-ARCHITECTURE.md` — the product boundary, workbook contract, and renderer rules.
- `docs/03-ROADMAP.md` — the commit-sized work items and gates.
- `docs/04-TEST-STRATEGY.md` — what to test and how.
- `docs/05-GIT-QUALITY.md` — commit, branch, and review rules.
- `docs/06-LLM-PROTOCOL.md` — anti-drift and anti-hallucination controls.
- `docs/07-GANTT-ENTITY-GUIDE.md` — geometry, style, label, and z-order contract for any visual work.
- `docs/STATUS.md` — current work item, recent commits, environment, and known limitations.
- `docs/work-items/<ID>.md` — the active work item.
- `docs/REVIEW-CHECKLIST.md` — the human reviewer's five-minute checklist.

## Non-negotiable rules

These are restated from the kit. Treat them as hard constraints.

- **Core boundary**: `GanttCreator.Core` has no Office, Excel-DNA, SkiaSharp, clipboard, filesystem-dialog, or UI dependency. Verified by an architecture test.
- **One sheet**: exactly one visible user worksheet plus one `_GanttCreatorConfig` `xlSheetVeryHidden` worksheet. Never another helper sheet.
- **Refresh-only render**: normal worksheet, Type, colour, label-position, and selection changes never render. Only the explicit Refresh command rebuilds the chart.
- **Evidence before claims**: do not invent Excel-DNA callbacks, COM members, RibbonX attributes, enum values, SkiaSharp APIs, package versions, test results, or user decisions. If a claim cannot cite a test or a primary source, it is `unknown` until evidence exists.
- **No loops**: two-attempt bounded retry. The third occurrence stops the agent and reports the diff, evidence, and one question.
- **Offline only**: the built product must not need the network. No telemetry, web fonts, cloud API, online licence check, or hidden fallback.

## Workflow

1. Read the active work item. Restate outcome, exclusions, acceptance tests, and unknowns.
2. Make the smallest coherent change that proves one behaviour.
3. Add or change tests in the same commit as behaviour.
4. Run `pwsh ./scripts/verify-quick.ps1` after each meaningful edit; `pwsh ./scripts/verify.ps1` before requesting review.
5. Update `docs/STATUS.md` with the result and any newly recorded known limitations.
6. Do not commit Office temporary files, build output, test results, customer images, or local logs.
7. Do not amend, rebase, force-push, tag, or publish unless explicitly asked.
