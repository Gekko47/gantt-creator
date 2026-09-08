---
name: 04-TEST-STRATEGY
description: Test layers, per-project coverage thresholds, the Core/raster/Office test catalogues, golden-image process, Office integration harness, fault injection, flaky-test policy, and PR test evidence. Trigger on writing or running tests, coverage, or golden-image questions.
---

Test layers, per-project coverage policy, the Core/raster/Office test
catalogues, golden-image process, and the flaky-test policy.

Do not get wrong:
- "Cover the whole codebase" means every production component has an
  appropriate automated or explicitly recorded host test — it does
  NOT mean chasing a 100% line-coverage number.
- Golden image updates require human review, a stated reason, and a
  dedicated commit — never bundled with unrelated behaviour changes.
- Office integration tests are tagged `OfficeIntegration`, serialize
  access, start from clean fixtures, and clean up in `finally`.
- Test observable behaviour, not private method implementation; Core
  and scene/layout tests must be deterministic and parallel-safe.

---

## Where to read more

- Full reference: `docs/04-TEST-STRATEGY.md` (this is the canonical source; there is no separate copy under .cline/skills/)
