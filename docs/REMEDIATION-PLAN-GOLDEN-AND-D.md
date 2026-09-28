# Remediation implementation plan — Phase 7 (golden regeneration) and Commit D

Follows `docs/REMEDIATION-PLAN-A.md`, whose Commit A landed as `4537b86`.

## Outcome

| Item | Status | Commit |
| --- | --- | --- |
| Phase 7 — golden regeneration | **Done, approved and committed** | `4278a49` |
| Commit D — `ApplyZOrder` ownership preflight | **Done** | `e5211ac` |

Both plans were written before execution, as required. Two findings during Phase 7 and
two during Commit D are recorded below because neither was anticipated by the plan
and both changed what was built.

### Phase 7 findings not anticipated by the plan

- **A second change rode along with the approved delta.** 21 primitives moved from
  `LaneOrder` *n* to *n*+1, because the `Splitter` now occupies lane 0. This was
  found by a normalised field-by-field comparison rather than by reading the diff, and
  it is intended — but it was not in the delta that was approved, so it is stated here
  and in STATUS rather than absorbed silently.
- **The plan assumed the harness wrote the source file; it writes the build output
  copy.** The regenerated file is copied to `tests/golden/scene/reference-scene.json`
  explicitly, and the harness is removed before the commit either way.

### Commit D findings not anticipated by the plan

- **The first three refusal tests were vacuous.** They asserted against
  `writer.ZOrderCommands`, which is wired only into shapes `AddShape` creates; a test
  built from seeded shapes therefore observed zero commands regardless of behaviour.
  Moving the ownership check into the mutation loop left all three *passing*. The
  mutation check is what caught this, and `NewShape` gained an optional `onZOrder`
  callback mirroring the existing `onDelete` convention.
- **The plan's "unowned same-name *later* shape" test could not work as written.** The
  approved pass walks the list in **reverse**, so a per-shape check inside the mutation
  loop processes the *last* entry first and refuses before issuing anything — the test
  would pass whether or not the check is preflighted. The atomicity test therefore puts
  the unowned shape in the **middle** of a three-element list, where a per-shape check
  has already moved a shape by the time it refuses.

With the ownership test moved into the mutation loop, **5 tests fail**; with it
restored, all 372 pass. Non-vacuity is proven, not asserted.

---

## Phase 7 — regenerate the golden scene snapshot (approved 2026-09-28)

### Why

Commit A made the fixture's `Splitter` (row 2) and `Spacer` (row 18) reach lane
layout for the first time. Both rows already existed in
`tests/fixtures/reference-gantt.json` and the old `SceneBuilder` discarded them, so
the committed baseline had never contained §10 geometry. The delta was presented for
human review and **approved** before any file was regenerated.

### Approved delta (55 → 59 primitives; chart bounds unchanged at 1120x394)

- **Added:** `…a1:splitter-band` (layer 30, y=110, h=18, w=1120),
  `…a1:splitter-band:top` and `…a1:splitter-band:bottom` (layer 80),
  `…a1:splitter-label` (layer 70). Nothing removed.
- **Moved by exactly +18pt** (= `SplitterHeightPt`): bars `b1` 113→131, `b2` 123→141,
  `b3` 133→151, `d1` 167→185, `e1` 177→195, `e2` 187→205, critical overlay `c1`
  113→131, milestone markers `f1`–`f4` 201/211/221/231 → 219/229/239/249.
- **Arithmetic check:** the `Splitter` is the fixture's first row, so every later lane
  is displaced once; the `Spacer` is the last row, so its 9pt displaces nothing. A
  displacement that was not a whole number of `SplitterHeightPt` would indicate a
  geometry defect, not an intended change.

### Steps

1. Write this plan (traceability before execution).
2. Regenerate `tests/golden/scene/reference-scene.json` through a **temporary** test
   harness, because the repository has no sanctioned regeneration script and the
   comparison must stay a byte comparison — the harness writes the file, it does not
   loosen the assertion.
3. **Verify the regenerated file is the approved delta and nothing more:** primitive
   count 55 → 59, chart bounds unchanged, four additions, seven-plus-four movements,
   every movement exactly +18, no other change.
4. Remove the temporary harness. It must not survive into the commit.
5. Run the Core suite: `The_built_fixture_scene_matches_the_committed_golden_snapshot`
   must now pass, and the suite must be fully green.
6. Confirm the other determinism pins still hold — the shuffle test
   (`Shuffled_fixture_rows_produce_an_identical_scene`) and the round-trip test
   (`The_golden_snapshot_round_trips_through_deserialization`) are the two that would
   catch a baseline that is merely self-consistent rather than correct.
7. Update STATUS: the golden row moves from "awaiting approval" to regenerated, with
   this commit's hash and the reason.
8. Commit alone, per the D4 policy: `docs:`/`test:` scope, no behaviour change.

### Explicitly not in this commit

No production code. The point of the D4 policy is that a golden regeneration is
reviewable in isolation, so bundling any behavioural change with it would defeat the
review.

## Commit D — `ApplyZOrder` ownership preflight (R4.7 blocker)

### Why

`ExcelShapeWriter.ApplyZOrder` resolves each named shape and calls `ZOrder` with **no
ownership check**, while `Update`, `Delete`, and `ListOwned` all require
`CarriesOwnershipTagFor`. A user-drawn shape that happens to share a requested
primitive ID would therefore be reordered by a refresh, even though the same shape is
correctly preserved by every other mutation path. This breaks the entity guide's
"refresh touches only shapes whose alternative text carries a valid tag".

The preflight for *existence* already exists: the method resolves every shape before
issuing the first command, so a missing shape cannot leave a partial order. The
ownership filter is simply missing from that same loop.

### Steps

1. Write this plan (done).
2. Add the ownership filter to the existing preflight loop, so an unowned shape
   refuses with **zero** `ZOrder` calls. Use the same `CarriesOwnershipTagFor`
   predicate and the same `NotFound` refusal `Update`/`Delete` already use, rather
   than inventing a new reason.
3. Do **not** touch the z-order algorithm. The single reverse-order
   `msoSendToBack` pass is an approved decision (ADR-0022-era review, and the
   entity guide's z-order contract); `The_z_order_pass_sends_each_shape_to_the_back
   _in_reverse_so_the_final_order_matches_the_scene` pins it, including a
   `DoesNotContain("BringToFront")` assertion.
4. Tests, all in `ExcelShapeWriterTests`:
   - unowned same-name shape **first** in the list → refusal, zero `ZOrder` calls;
   - unowned same-name shape **later** in the list → refusal, zero `ZOrder` calls
     (the atomicity case, and the one that would regress if the check were placed
     inside the mutation loop);
   - all owned → the existing deterministic sequence still asserted, unchanged;
   - a user shape not named by the scene is untouched (already covered by the
     existing preservation tests; re-assert rather than duplicate).
5. Non-vacuity: each refusal test must be mutation-checked by temporarily moving the
   check into the mutation loop and observing the first case start issuing calls.
6. Update `docs/STATUS.md`; add the R4.7 work-item note that the preflight
   obligation is discharged.
7. Commit separately from Phase 7.

### Explicitly not in this commit

No change to the approved ordering algorithm; no change to `Create`, `Update`, or
`Delete`; no reconciliation logic (that is R4.7 proper). This is the small
independent blocker only.

## Still deferred after both

| Commit | Work | Why not now |
| --- | --- | --- |
| B | Per-row `PanelCellGrid`; one panel-bounds authority; panel rows from source rows rather than lane placements | Deliberately breaking; needs its own approval round |
| C | `IPanelGridMeasurementPort` per-row measurement; mixed heights; the read-only protection policy | Blocked on B, and the protection half is an ADR-0008 decision |
| E | R4.6 style completeness | Separate R4 row |
| F | R4.7 reconciliation | Blocked on E |
