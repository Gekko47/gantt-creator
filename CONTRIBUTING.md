# Contributing to Gantt Creator

Thank you for your interest in contributing. This document covers how to propose, implement, and land changes.

## Prerequisites

- Windows 11 (current), x64
- .NET 10 SDK, x64
- 64-bit Microsoft 365 desktop Excel and PowerPoint (for Office integration tests)
- Visual Studio 2022 17.14+ or VS Code with the C# Dev Kit
- PowerShell 7 and Git for Windows

See `docs/01-ENVIRONMENT.md` for the full setup checklist.

## Development workflow

1. **Pick a work item** from `docs/03-ROADMAP.md` (Phase 0 is complete; current work is post-Phase-0 hardening). Create a work-item file from `docs/work-items/TEMPLATE.md` before implementing.
2. **Read the relevant kit docs**: `AGENTS.md` (always-on contract), `docs/02-ARCHITECTURE.md` (product boundary), `docs/04-TEST-STRATEGY.md` (test layers), and the affected `docs/07-GANTT-ENTITY-GUIDE.md` sections for visual work.
3. **Restate the boundary**: outcome, files likely to change, exclusions, acceptance tests, and unknowns.
4. **Make the smallest coherent change** that proves one behaviour. Add or change tests in the same commit.
5. **Run the gates**:
   - `pwsh ./scripts/verify-quick.ps1` after each meaningful edit
   - `pwsh ./scripts/verify.ps1` before requesting review
6. **One concern per commit**. Conventional Commit prefixes (`feat:`, `fix:`, `test:`, `refactor:`, `docs:`, `build:`, `chore:`).
7. **Update control documents**: `docs/STATUS.md` after every commit; `docs/DECISIONS.md` and `docs/KNOWN-LIMITATIONS.md` only when necessary.
8. **Do not** commit Office temporary files, build output, test results, exported customer images, or local logs.

## Architecture rules (non-negotiable)

- `GanttCreator.Core` has **no** Office, Excel-DNA, SkiaSharp, clipboard, filesystem-dialog, or UI dependency. Enforced by `tests/GanttCreator.Architecture.Tests/CoreBoundaryTests.cs`.
- Exactly one visible user worksheet plus one `_GanttCreatorConfig` `xlSheetVeryHidden` worksheet. Never another helper sheet.
- Normal worksheet, Type, colour, label-position, and selection changes never render. Only the explicit Refresh command rebuilds the chart.
- Multiple events on one lane, overlapping planned/actual events, milestones, critical intervals, and labelled full-height delineators are first-class domain cases.

## Code style

- Nullable reference types, implicit usings, analyzers, and warnings-as-errors are enabled (see `Directory.Build.props`).
- Prefer immutable records/value objects in Core. Use `DateOnly` for date-only schedule data.
- Public APIs need XML documentation when the contract is not obvious.
- xUnit test method names use snake_case (e.g., `Rotation_at_cap_never_deletes_active_log`).

## Review process

1. Open a pull request against `main`.
2. `pwsh ./scripts/verify.ps1` must PASS before review.
3. One human approval required.
4. Address review findings with new commits; do not amend or rebase shared history.

## Scope-change protocol

If evidence invalidates a fixed design: stop, preserve the spike outside production code, write an ADR in `docs/adr/`, obtain human approval, then update architecture, roadmap, tests, and agent rules together.

## Licence

By contributing, you agree that your contributions will be licensed under the MIT License (`LICENSE`).
