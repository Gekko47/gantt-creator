---
name: 04-TEST-STRATEGY
description: Use this skill when the conversation is about test layers, the per-project coverage thresholds, the Core test catalogue, the Excel shape contract tests, the raster tests and golden images, the Office integration harness, the feature-to-test traceability map, fault injection, the flaky-test policy, or the pull-request test-evidence block.
---

# Test strategy
## Test objective
Testing must prove the domain and rendering logic without Office, then prove the thin Office adapters on a controlled Windows machine. “Cover the whole codebase” means every production component has an appropriate automated or explicitly recorded host test; it does not mean pursuing a misleading 100% line-coverage number.
## Test layers
| Layer | Runs where | Purpose | Release expectation |
| --- | --- | --- | --- |
| Core unit/property | any `.NET 10` runner | validation, date math, lanes, scene geometry, ordering | every commit |
| Raster unit/golden | pinned Windows runner | exact pixels, fonts, crop, PNG metadata | every PR |
| Office contract | Windows, no live Office where possible | calls and state transitions through fakes/adapters | every commit/PR |
| Architecture | build runner | dependency direction, naming/tagging constraints | every commit |
| Office integration | dedicated Windows machine with supported Office | real Excel-DNA, Excel shapes, clipboard, PowerPoint | phase exit and release; nightly if available |
| Manual visual/accessibility | controlled reference machine | human judgement and Office UI behaviour | relevant phase exit/release |
| Performance/reliability | reference machine | budgets, repeated transfers, leaks/hangs | phase exit/release |
## Coverage policy
Initial thresholds are deliberately strict in deterministic code and realistic at external boundaries:
| Project | Line | Branch | Additional gate |
| --- | ---: | ---: | --- |
| `GanttCreator.Core` | 95% | 90% | mutation score at least 80% on changed Core code |
| `GanttCreator.Raster` | 90% | 85% | approved representative golden images |
| `GanttCreator.Office` | 80% | 70% | Office contract matrix plus live-host tests |
| `GanttCreator.AddIn` | 75% | 65% | Ribbon XML/callback and command-boundary tests |
Exclude only generated interop/build files and trivial assembly metadata. Never exclude a file because it is difficult to test. A threshold change requires an ADR or explicit pull-request approval with evidence.
Coverage is a floor, not proof. Review changed lines, branches, fault paths, and assertions. Run mutation testing on Core weekly or before release; do not burden every small local commit with a full mutation run.
## Core test catalogue
### Validation and schema
- required/missing/duplicate columns;
- blank, duplicate, and stable IDs;
- supported/unknown event type;
- span with blank date, reverse dates, and one-day range;
- point event with one or conflicting dates;
- whitespace and Unicode text;
- duplicate stack positions and missing lane;
- events entirely/partly outside plot range;
- all errors returned in stable row/field order;
- workbook 1900 date system; 1904 either supported and tested or rejected explicitly;
- date values around Excel's invalid/edge serial values.
### Time and geometry
- first/last day and inclusive-finish policy;
- leap day, month/year boundary, very long programme;
- identical start/finish, zero plot width, invalid values;
- date-to-point and point-to-date boundary behaviour;
- lane order independent of input enumeration order;
- `StackIndex` layout for one, two, and many events;
- clipping at left/right/top/bottom;
- labels at each placement, overflow, and collision policy;
- stable primitive IDs and z-order;
- no NaN, Infinity, or negative-size primitive;
- deterministic scene serialization from shuffled input and repeated runs.
Use generated/property tests for invariants such as: valid dates map monotonically; clipping never expands bounds; render order is stable; all emitted geometry remains inside allowed scene bounds except intentional label overflow.
## Numeric, configuration, and suppression tests
These apply across all layers and are required in addition to the
domain-specific catalogue above. See also `docs/08-TEST-CHECKLIST.md`.
### Numeric code
- Culture-roundtrip: tests must set `CurrentCulture` to a comma-decimal
  culture (e.g. `de-DE`) before invoking any public method that parses,
  formats, or converts numeric values, and assert the exact output /
  parsed value. Production methods must not be required to modify ambient
  culture.
- Non-finite inputs: `double.TryParse` consumers must test `NaN`,
  `+Infinity`, `-Infinity`, and overflowed exponents (`"1e999"`).
- Every `double`/`float` to `int` cast must be guarded by `IsFinite` and a
  range check; tests cover each violation.
### Configuration and schema files
- Machine-readable config (JSON, YAML, RibbonX) must have an architecture
  test validating structure on every commit. For `.vscode/tasks.json`:
  assert the VS Code `{ version, tasks }` object form, that every `dependsOn`
  label resolves to a task, and that dependency ordering is correct.
### Build pipeline traceability
- Every test that reads from `bin/` or `publish/` must be traceable to a
  step in `verify-quick.ps1` / `verify.ps1` that produces that artifact. The
  guarantee is documented in the work item, not assumed by the test.
### NoWarn scope
- Test-only suppressions (CA1707, IDE0011) must live in
  `tests/Directory.Build.props`, never in the root `Directory.Build.props`.
  An architecture test asserts this on every commit.
## Excel shape contract tests
The renderer talks to narrow fake adapters in most tests. Assert the observable Office operation sequence and values:
- exact shape type, name, ownership tag, and stable scene ID;
- point geometry within the documented COM tolerance;
- fill/stroke/font/alignment/pattern properties;

---

## Where to read more

- Canonical source: `docs/04-TEST-STRATEGY.md`
- Always-on rule: `.clinerules/04-TEST-STRATEGY.md`
- Full reference: `./references.md` in this directory
