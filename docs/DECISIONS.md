# Decisions log (ADR index)

> Lightweight Architecture Decision Records. One decision per file in `docs/adr/`.
> This file is the index. The file body in each ADR is the durable record.

| ID | Title | Status | Date | File |
| --- | --- | --- | --- | --- |
| ADR-0000 | Adopt the Gantt Creator development kit as the source of truth | Accepted | 2026-09-04 | [`adr/0000-adopt-kit.md`](adr/0000-adopt-kit.md) |
| ADR-0001 | Self-hosted Windows runner as the Office-integration host | Accepted | 2026-09-04 | [`adr/0001-self-hosted-office-runner.md`](adr/0001-self-hosted-office-runner.md) |
| ADR-0002 | Commit `.clinerules/*` and `AGENTS.md` as canonical sources | Accepted | 2026-09-04 | [`adr/0002-commit-clinerules-and-agents.md`](adr/0002-commit-clinerules-and-agents.md) |
| ADR-0003 | Visual Studio 2026 is the verification IDE; Visual Studio 2022 baseline is relaxed | Accepted | 2026-09-04 | [`adr/0003-visual-studio-2026.md`](adr/0003-visual-studio-2026.md) |
| ADR-0004 | Pin Excel-DNA packages to AddIn 1.9.0, Integration 1.9.0, Interop 16.0.0 | Accepted | 2026-09-04 | [`adr/0004-exceldna-package-versions.md`](adr/0004-exceldna-package-versions.md) |
| ADR-0005 | Document the actual host as Windows 11 25H2 (build 26200) | Accepted | 2026-09-04 | [`adr/0005-windows-host-build-number.md`](adr/0005-windows-host-build-number.md) |
| ADR-0006 | Reject the 1904 workbook date system on read | Accepted | 2026-09-19 | [`adr/0006-reject-1904-date-system.md`](adr/0006-reject-1904-date-system.md) |
| ADR-0007 | Configuration-sheet catalogues: Excel Tables, `tblGanttSettings`, and the catalogue hash | Accepted | 2026-09-21 | [`adr/0007-config-sheet-catalogues.md`](adr/0007-config-sheet-catalogues.md) |
| ADR-0008 | Destructive-command policy: no undo, warned confirmation | Accepted | 2026-09-22 | [`adr/0008-destructive-command-policy.md`](adr/0008-destructive-command-policy.md) |
| ADR-0009 | Named-style capability columns | Accepted | 2026-09-23 | [`adr/0009-style-capability-schema.md`](adr/0009-style-capability-schema.md) |
| ADR-0010 | First-live-slice sequencing exception | Accepted | 2026-09-23 | [`adr/0010-first-live-slice-sequencing.md`](adr/0010-first-live-slice-sequencing.md) |
| ADR-0011 | Safe configuration and identity repair | Accepted | 2026-09-24 | [`adr/0011-config-safe-repair-policy.md`](adr/0011-config-safe-repair-policy.md) |
| ADR-0012 | Derive effective stack index in Core | Accepted | 2026-09-25 | [`adr/0012-derived-effective-stack-index.md`](adr/0012-derived-effective-stack-index.md) |
| ADR-0013 | Chart-level scene ownership | Accepted | 2026-09-25 | [`adr/0013-chart-level-scene-ownership.md`](adr/0013-chart-level-scene-ownership.md) |
| ADR-0014 | Versioned frame and title settings | Accepted | 2026-09-25 | [`adr/0014-frame-band-settings.md`](adr/0014-frame-band-settings.md) |

## ADR template

```markdown
# ADR-<NNNN> — <title>

- **Status**: Proposed | Accepted | Superseded by ADR-<NNNN>
- **Date**: YYYY-MM-DD
- **Context**: the forces at play, the constraint, the question.
- **Decision**: what we decided.
- **Consequences**: trade-offs, what becomes easier, what becomes harder.
- **Alternatives considered**: what we rejected, and why.
```
