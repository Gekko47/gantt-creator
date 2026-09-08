---
name: 02-ARCHITECTURE
description: Product boundary; one-visible-sheet + _GanttCreatorConfig contracts; scene-first rendering and renderer equivalence; COM ownership; binding C# and architecture conventions (nullable/analyzers, immutability, DateOnly, geometry rounding, DI boundaries, save/restore, XML doc). Trigger on architecture, worksheet-schema, COM, or C#-convention questions.
---

Product boundary, worksheet contracts, scene-first rendering model,
and COM/error-handling rules for the add-in.

Do not get wrong:
- Exactly one visible worksheet plus one `xlSheetVeryHidden`
  `_GanttCreatorConfig` worksheet — never a second helper sheet, and
  never schedule rows, shapes, or logs on the VeryHidden sheet.
- `verify-quick.ps1`/`verify.ps1` are the only authoritative producers
  of `bin/`, `publish/`, and `coverage/` — a test that reads one of
  these artifacts without naming the producing step is a drift risk
  (see docs/04-TEST-STRATEGY.md "Build pipeline traceability").
- This is not a CPM scheduling engine; it must never imply it
  calculates contractual entitlement or delay causation.
- The "C# and architecture conventions" section at the end of this doc
  is binding for every C# change: nullable/analyzers on, immutable Core
  types with `DateOnly`, point-based geometry with one rounding policy,
  injected clock/dialog/clipboard/Office/logging boundaries, no
  `dynamic`, Excel state save/restore in `try/finally`.

---

## Where to read more

- Full reference: `docs/02-ARCHITECTURE.md` (this is the canonical source; there is no separate copy under .cline/skills/)
