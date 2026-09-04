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

---

## Where to read more

- Canonical source: `docs/04-TEST-STRATEGY.md`
- Always-on rule: `.clinerules/04-TEST-STRATEGY.md`
- Full reference: `./references.md` in this directory
