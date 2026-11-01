---
name: 02-architecture
description: Product boundary and binding architecture rules for the Gantt Creator add-in — the one-visible-worksheet plus xlSheetVeryHidden _GanttCreatorConfig contract, scene-first rendering and renderer equivalence, build-pipeline artifact traceability, COM ownership and application-state scope, Excel/PowerPoint/PNG/ribbon boundaries, and the C# conventions (nullable/analyzers, immutability, DateOnly, geometry rounding, DI boundaries, XML docs). Part of the gantt-creator index. Trigger on architecture, worksheet-schema, COM, or C#-convention questions.
---

# Architecture and C# conventions

Mirror of [`.cline/skills/02-architecture/SKILL.md`](../../../.cline/skills/02-architecture/SKILL.md),
generated from [`docs/02-ARCHITECTURE.md`](../../../docs/02-ARCHITECTURE.md). The canonical doc
wins on any disagreement; this summary is not the contract.

## Index

- Index first: `gantt-creator-contract`; `00-ponytail` before any code change.
- Pair with `07-entity-guide` for anything visual, and `04-test-strategy` for the test layer
  the change belongs to.

## Do not get wrong

- Exactly one visible worksheet plus one `xlSheetVeryHidden` `_GanttCreatorConfig` for the
  adopted-blank target, or two visible worksheets for the create path (the preserved original
  plus the new `Gantt Data` sheet) plus one `_GanttCreatorConfig` — never a second helper
  sheet, never schedule rows, shapes, or logs on the VeryHidden sheet, and never hide an
  existing worksheet during initialization.
- `verify-quick.ps1` / `verify.ps1` are the only authoritative producers of `bin/`,
  `publish/`, and `coverage/` — a test that reads one of those artifacts without naming the
  producing step is a drift risk (see `docs/04-TEST-STRATEGY.md`, "Build pipeline
  traceability").
- This is not a CPM scheduling engine; it must never imply it calculates contractual
  entitlement or delay causation.
- The closing "C# and architecture conventions" section is binding for every C# change:
  nullable/analyzers on, immutable Core types with `DateOnly`, point-based geometry with one
  rounding policy, injected clock/dialog/clipboard/Office/logging boundaries, no `dynamic`,
  Excel state save/restore in `try/finally`.
- COM: Excel/PowerPoint calls on the required STA/main thread; no `Task.Run` around COM; no
  chained proxy expressions; every proxy held locally and released by one ownership helper.

## Tools here

`codebase_search` / `search_files` (find a port, adapter, or token) · `read_file` ·
`execute_command` (`dotnet build -warnaserror`, `dotnet test
tests/GanttCreator.Architecture.Tests`, `git grep` for pattern violations such as
`Thread.Sleep` or `Task.Run` around COM).
