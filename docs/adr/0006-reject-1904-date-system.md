# ADR-0006 — Reject the 1904 workbook date system on read

- **Status**: Accepted
- **Date**: 2026-09-19
- **Context**: Checklist F (`docs/08-TEST-CHECKLIST.md`) requires the 1900 date system and forces a choice for 1904: support-and-test or reject explicitly with an ADR. R2.4 reads `Start`/`Finish` from `Range.Value2` serials, whose epoch depends on `Workbook.Date1904`. Supporting both epochs doubles date-conversion tests and leaks migration policy into the first read path.
- **Decision**: R2.4 requires the 1900 date system. When `Workbook.Date1904` is `true`, `ExcelGanttTableReader` returns the typed refusal `DateSystemUnsupported` before touching the table; nothing is mutated. The Core converter takes the flag as a primitive (`bool date1904`) so the refusal branch is unit-testable without Office. Schema stays v1.
- **Consequences**: 1904 workbooks get a clear actionable refusal instead of silently shifted dates. A future decision may add 1904 support (+1462-day offset) with its own tests and migration policy; that is R2.10 scope, not R2.4.
- **Alternatives considered**: Supporting 1904 via offset in R2.4 (rejected — doubles test surface and pre-empts R2.10 migration design). Ignoring the flag (rejected — silently shifts every date by 4 years + 1 day).
