# ADR-0010 — First-live-slice sequencing exception

- **Status**: Accepted
- **Date**: 2026-09-23
- **Context**: The roadmap normally executes rows in table order. The R2 findings review shows that valid authoring work does not require R2.10 repair UI or R3.13 mutation-test infrastructure, while both currently delay the first usable live Gantt. R2.10 and R3.13 remain mandatory and may expose product or dependency decisions.
- **Decision**: After R2.9, execute R3.1-R3.12 and R3.14, then R4.1-R4.9. Execute R2.10 and R3.13 immediately after the R4.9 first-live-slice demonstration, then R4.10. Neither deferred row is removed, relaxed, or allowed to block valid read/validate/scene/render operation. R2.10 still requires its separately approved safe-repair ADR before code; R3.13 still requires mutation-tool approval.
- **Consequences**: The first live product arrives before recovery UI and test-tool infrastructure. Phase labels remain historical; the manifest and roadmap notes carry the executable order. CI and local verification remain mandatory for every row.
- **Alternatives considered**: Keep table order (rejected: delays the first working product without reducing required scope); remove R2.10 or R3.13 (rejected: both remain release obligations); implement partial versions early (rejected: creates provisional contracts).
