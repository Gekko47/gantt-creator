# ADR-0066 — Deleting staging shapes after `ShapeRange.Copy` preserves the Office drawing clipboard format, so immediate cleanup stands

- **Status:** Accepted (owner-approved 2026-11-18; satisfies R6.7 D3 and unblocks R6.5/R6.6)
- **Date:** 2026-11-18
- **Related:** R6.7 (the spike), [`docs/02-ARCHITECTURE.md`](../02-ARCHITECTURE.md)
  "Editable export composition" step 7 and its fallback paragraph, R6.5 (staging
  lifecycle), R6.6 (clipboard copy), R7.1/R7.3 (PowerPoint adapters and paste
  verification), R10.1/R10.5 (supported-Office matrix and full-suite re-runs)

## Context

The editable export stages native shapes, groups them, calls `ShapeRange.Copy`, and
then deletes the staging shapes in `finally` (architecture step 7). The open question
was whether deleting the source shapes after the copy invalidates the delayed clipboard
rendering that PowerPoint's `PasteSpecial(ppPasteShape)` and an Excel paste depend on.
If it did, the approved fallback is a short-lived staging workbook retained until the
copy materialises — never a permanent helper worksheet.

R6.7 answered the question with a measured spike before any composition row was built.
The spike drives real Excel and PowerPoint on the reference host, records the clipboard
format set through Win32 `EnumClipboardFormats` and `GetClipboardFormatNameW`, and then
materialises the clipboard into both targets. Each delete cycle is paired with a
no-delete control cycle so a failure is attributable to the deletion. Two fixed
measurement windows (2 s and 10 s) expose spontaneous invalidation; because delayed
rendering is serviced only when a consumer requests the format, the paste assertions —
not the windows — are the primary measurement.

## Evidence (measured 2026-11-18)

**Hosts.** Excel 16.0 build 20527; PowerPoint 16.0.20527.20196 (x64, O365HomePremRetail).
This is the single-build evidence basis: the supported-Office matrix does not exist yet
(R10.1 owns it), and R10.1/R10.5 must re-run this check on every build they add.

**Clipboard format set after the copy** (identical immediately, after 2 s, and after
10 s, in every cycle and arm):

```text
DataObject, Preferred DropEffect, InShellDragLoop, Excel 2007 Internal Shape,
Art::GVML ClipFormat, image/svg+xml, PNG, JFIF, GIF, CF_BITMAP, CF_ENHMETAFILE,
CF_METAFILEPICT, Ole Private Data, CF_DIB, CF_DIBV5
```

`GetClipboardSequenceNumber` was unchanged between the immediate read and both delayed
reads in every cycle (for example 548 → 548 → 548), so the deletion did not cause the
clipboard content to be replaced or withdrawn.

**Staging fixture.** Four named shapes (three `msoShapeRectangle`, one text box) at a
fixed origin on a test-owned worksheet, grouped into one group of four. The same
four-shape composition is the minimum visual fixture the entity guide defines; the
staging prefix is deliberately distinct from the production `GanttCreator.Owned.v1` tag.

| Arm | Cycles | Clipboard format survived deletion | Excel paste | PowerPoint `ppPasteShape` | Ungroup |
| --- | --- | --- | --- | --- | --- |
| Delete, Excel target | 3 of 3 | Yes (all 15 formats, both windows) | 1 group, 4 children, 340 × 190 pt | — | 4 editable children |
| Delete, PowerPoint target | 3 of 3 | Yes (all 15 formats, both windows) | — | range of 5: the group plus 4 flattened children; group has 4 children | 4 editable children |
| Control, no delete, Excel target | 1 of 1 | Yes | 1 group, 4 children, 340 × 190 pt | — | 4 editable children |

Every row is from a green run of
[`CopyAfterDeleteSpikeIntegrationTests`](../../tests/GanttCreator.Office.IntegrationTests/CopyAfterDeleteSpikeIntegrationTests.cs);
the raw JSON-lines evidence was written to the ignored office-evidence folder and the
table above is its committed summary.

## Decision

### D1 — The simple delete path stands

Staging shapes may be deleted immediately after `ShapeRange.Copy` in `finally`. The
short-lived staging-workbook fallback is **not** required by this evidence and is not
built. R6.5's lifecycle is the immediate-cleanup design.

### D2 — R6.6 verifies materialisation, not only format presence

The format a shape copy leaves on this build is **`Excel 2007 Internal Shape`**, with
`Art::GVML ClipFormat` alongside it. The name "Office Drawing Shape Format" — the
candidate the plan expected — was not present. R6.6's clipboard contract therefore
asserts that the copy leaves the measured drawing format **and** that a real consumer
paste materialises the group; a non-throwing `Copy` call or a format-name guess is not
evidence.

### D3 — R7.1/R7.3 must expect the group inside a larger pasted range

PowerPoint's `PasteSpecial(ppPasteShape)` of an Excel group returns a range of five
shapes: the group (`msoGroup`, four children) plus the four flattened children. The
pasted range is not a single-item range, and the group must be located within it. R7.1's
adapter and R7.3's verification must not assume `ShapeRange.Count == 1`, and cleanup
must delete the group rather than every item in the range (the children are covered by
the group).

### D4 — The fallback remains the approved remedy on divergent builds

If a future supported build (R10.1/R10.5) shows a format that fails to materialise
after deletion, the short-lived staging-workbook fallback in architecture step 7 is the
approved remedy, decided from that build's evidence. This ADR does not generalise its
single-build result to unmeasured builds.

### D5 — Evidence basis is stated, including its simplification

The spike staged on a fixture-owned worksheet, not on a user's live sheet beside real
content. The clipboard mechanics under test are host-level, not worksheet-level, but the
simplification is recorded rather than left implicit. The delay windows are measurement
parameters, not synchronization; no retry-until-green logic exists.

## Consequences

- **R6.5** implements immediate staging cleanup with fault-injection tests; no workbook
  retention path is added.
- **R6.6** pins the measured drawing format and a paste-materialisation assertion.
- **R6.1–R6.4/R6.8** proceed on the simple delete path; none of them inherits a
  staging-workbook dependency.
- **R7.1/R7.3** carry the five-item pasted-range behaviour into the production adapter's
  verification and cleanup rules.
- **R10.1/R10.5** must re-run the copy-after-delete check on every supported build and
  record the result; a divergent build reopens this ADR rather than being averaged.
- Approving this ADR (R6.7 D3) is the gate that unblocks the composition rows.
