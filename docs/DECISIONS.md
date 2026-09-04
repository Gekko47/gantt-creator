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
