---
name: 04-test-strategy
description: Test layers, per-project coverage policy, the Core/raster/Office test catalogues, golden-image process, Office integration harness, fault injection, gate-integrity and step-parity tests, flaky-test policy, and PR test evidence for the Gantt Creator repository. Part of the gantt-creator index. Trigger on writing or running tests, coverage, golden-image, or Office-harness questions.
---

# Test strategy

Mirror of [`.cline/skills/04-test-strategy/SKILL.md`](../../../.cline/skills/04-test-strategy/SKILL.md),
generated from [`docs/04-TEST-STRATEGY.md`](../../../docs/04-TEST-STRATEGY.md). The canonical doc
wins on any disagreement.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- Pair with `02-architecture` for layer boundaries and `05-git-quality` for what the commit
  and PR must show.

## Do not get wrong

- "Cover the whole codebase" means every production component has an appropriate automated or
  explicitly recorded host test — it does NOT mean chasing a 100% line-coverage number.
- Golden image / golden scene updates require human review, a stated reason, and a dedicated
  commit — never bundled with unrelated behaviour changes.
- Office integration tests are tagged `OfficeIntegration`, serialize access, start from clean
  fixtures, and clean up in `finally`. They are excluded from `verify-quick.ps1` by policy
  (ADR-0001) and are run by `verify-office.ps1` on the live host.
- Test observable behaviour, not private method implementation; Core and scene/layout tests
  must be deterministic and parallel-safe.
- Every new `if (bad) { error }` validator ships a positive test in the same commit — a test
  that constructs the bad input and asserts the error path fires, in the harness that will run
  it. For PowerShell the equivalent is a harness under `scripts/` invoked from
  `verify-quick.ps1`.
- No `Thread.Sleep` for synchronisation: poll a named observable condition with a deadline and
  useful timeout diagnostics. No arbitrary delays, no catch-and-ignore, no deleted tests.

## Tools here

`execute_command` (`dotnet test` per project, `dotnet build`, `pwsh -NoProfile -File
scripts/verify-quick.ps1`, `verify.ps1`, `verify-office.ps1`) · `codebase_search` /
`search_files` (trace an acceptance criterion to a test name) · `read_file`.

## Test projects

`GanttCreator.Core.Tests`, `GanttCreator.Raster.Tests`, `GanttCreator.Office.ContractTests`
(Moq PIA, no live Office), `GanttCreator.Office.IntegrationTests` (live host,
`OfficeIntegration`), `GanttCreator.AddIn.Tests`, `GanttCreator.Architecture.Tests`.
