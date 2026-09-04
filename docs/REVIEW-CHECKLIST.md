# Pull request review checklist

> Run by the human reviewer for every PR. Tick each item or write a one-line note. Five minutes per PR.

## Behaviour

- [ ] Acceptance criteria in `docs/work-items/<ID>.md` are met; no extra product behaviour appeared.
- [ ] Error / empty / boundary cases are visible in tests.
- [ ] The one-sheet rule is still true: the workbook has exactly one visible Gantt worksheet and one valid `_GanttCreatorConfig` VeryHidden worksheet.
- [ ] The offline rule is still true: no telemetry, web fonts, cloud API, online licence check, or hidden network fallback.

## Architecture

- [ ] `GanttCreator.Core` has no Office, Excel-DNA, SkiaSharp, clipboard, filesystem-dialog, or UI reference. (Verified by the architecture test in `tests/.../Architecture.Tests`.)
- [ ] Scene layout is not duplicated in an Office or raster renderer. Renderers consume the resolved scene.
- [ ] New COM calls are isolated behind an adapter and have a contract or integration test.
- [ ] New configuration has a schema version and a migration / default policy.
- [ ] `_GanttCreatorConfig` is the only helper worksheet; no second helper has appeared.

## Quality

- [ ] Assertions would fail for a plausible defect.
- [ ] No suppressed warning, reduced threshold, hidden retry, arbitrary `Thread.Sleep`, or catch-and-ignore.
- [ ] No unnecessary dependency or public API; any new package has an ADR or an explicit PR approval.
- [ ] Logs exclude user schedule content by default; any include-content path is opt-in with a reason.
- [ ] `scripts/verify-quick.ps1` and `scripts/verify.ps1` exit 0 on the PR commit (linked in the PR description).

## Readability

- [ ] Names make the normal path clear.
- [ ] Comments explain *why*, *invariants*, or *Office quirks* — not *what* the code does line by line.
- [ ] XML documentation describes real public contracts, not implementation details.
- [ ] No commented-out code, no stale TODOs, no generated noise.

## Visual / golden (only if a renderer or style changed)

- [ ] The Office integration or golden-image check actually ran; the evidence is linked in the PR.
- [ ] The renderer or style change is reflected in `docs/07-GANTT-ENTITY-GUIDE.md` and the entity tests.
- [ ] No `BringToFront` opportunism; z-order is applied deterministically.
