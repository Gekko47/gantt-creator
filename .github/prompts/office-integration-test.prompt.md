---
mode: agent
description: Add or fix an Office-integration test
---

# Office integration test

You are adding or fixing an Office-integration test for **${input:ID:work item or scenario, e.g. R1.1 AutoOpen logging}**.

## Preconditions

- The test is tagged `[Trait("Category", "OfficeIntegration")]` and is excluded from `scripts/verify.ps1`.
- The test is run by `scripts/verify-office.ps1` on a self-hosted Windows runner with a real licensed Microsoft 365 x64 install (per `docs/adr/0001-self-hosted-office-runner.md`).
- Record the actual Office version, build, channel, locale, display scale, and test run ID in the test output.

## Structure

- One test class per Office-host scenario. Serialized via a single xUnit collection.
- Start from a clean fixture file in `tests/fixtures/`. Copy to a unique temporary directory per run.
- Acquire the COM proxies explicitly. Capture each proxy in a local variable. Release in reverse order in `finally`.
- Do not terminate any Office process you did not start.
- Poll named observable state with a deadline. Never `Thread.Sleep` for sync.
- Delete temporary files in `finally`. Retain failed evidence (logs, screenshots) when the run fails.

## Required live tests (from `docs/04-TEST-STRATEGY.md`)

1. XLL loads, Ribbon XML is accepted, callbacks resolve, and shutdown completes.
2. Blank workbook initialises with one sheet and correct visible table/plot anchors.
3. Representative fixture renders expected owned shape types/counts/bounds/z-order.
4. Second refresh is idempotent and preserves an unowned sentinel shape/cell.
5. Save/reopen retains data/settings with one sheet.
6. Editable composition copies; pasted result is a group with editable children.
7. Temporary shapes are removed after success and injected failure.
8. `PasteSpecial(ppPasteShape)` returns a non-empty `ShapeRange` on the active slide.
9. PowerPoint transfer does not save, close, or kill user-owned content/application.
10. PNG export has exact size, aspect, crop, and density metadata.
11. Protected sheet, missing PowerPoint, busy clipboard, and unwritable path fail safely.
12. Twenty-five repeated refresh/copy/transfer cycles do not leave orphaned Office processes or steadily growing owned shapes.

## Before you hand back

1. The new test runs in the Office-integration suite and passes against the actual Office build on this machine.
2. `pwsh ./scripts/verify-quick.ps1` (which excludes OfficeIntegration by filter) still exits 0.
3. The PR description includes the actual Office build, locale, display scale, and the test run ID.
