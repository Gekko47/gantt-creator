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
| ADR-0015 | Label `Auto` cascade order and the widest-gap truncation fallback | Accepted | 2026-09-26 | [`adr/0015-label-cascade-and-widest-gap-fallback.md`](adr/0015-label-cascade-and-widest-gap-fallback.md) |
| ADR-0016 | Approved date display format for event dates | Accepted | 2026-09-26 | [`adr/0016-date-display-format-setting.md`](adr/0016-date-display-format-setting.md) |
| ADR-0017 | Shared scene ownership for deduplicated entities | Accepted | 2026-09-26 | [`adr/0017-shared-scene-ownership.md`](adr/0017-shared-scene-ownership.md) |
| ADR-0018 | Text alignment is its own type, not a label position | Accepted | 2026-09-26 | [`adr/0018-text-alignment-is-not-a-label-position.md`](adr/0018-text-alignment-is-not-a-label-position.md) |
| ADR-0019 | Shape ownership carrier is `AlternativeText`, with a bounded versioned tag | Accepted | 2026-09-27 | [`adr/0019-shape-ownership-carrier.md`](adr/0019-shape-ownership-carrier.md) |
| ADR-0020 | The application-state scope covers five settings, not six | Accepted | 2026-09-27 | [`adr/0020-application-state-scope-five-settings.md`](adr/0020-application-state-scope-five-settings.md) |
| ADR-0021 | Structural rows reach lane layout; "empty" means no visible entity | Accepted | 2026-09-28 | [`adr/0021-structural-rows-and-empty-scene.md`](adr/0021-structural-rows-and-empty-scene.md) |
| ADR-0022 | The §10 splitter band is a new scene builder on the existing style pipeline | Accepted | 2026-09-28 | [`adr/0022-splitter-band-builder.md`](adr/0022-splitter-band-builder.md) |
| ADR-0023 | The panel measures per row, and the panel's bounds are derived | Accepted | 2026-09-28 | [`adr/0023-panel-measures-per-row.md`](adr/0023-panel-measures-per-row.md) |
| ADR-0024 | Measure each row; a read-only measurement does not refuse a protected sheet | Accepted | 2026-09-28 | [`adr/0024-per-row-measurement-and-read-only-protection.md`](adr/0024-per-row-measurement-and-read-only-protection.md) |
| ADR-0030 | The live chart is anchored to the worksheet rows, not to page coordinates | Accepted | 2026-10-01 | [`adr/0030-live-chart-vertical-anchoring.md`](adr/0030-live-chart-vertical-anchoring.md) |
| ADR-0033 | Labels: no fill, no border, no margin, one row tall, and stacks are not collisions | Accepted | 2026-10-02 | [`adr/0033-label-appearance-and-stack-collision.md`](adr/0033-label-appearance-and-stack-collision.md) |
| ADR-0034 | Lanes are anchored to their worksheet rows, the content block is centred, and the critical interval is top-aligned | Accepted | 2026-10-02 | [`adr/0034-lane-anchoring-centring-critical-top-align.md`](adr/0034-lane-anchoring-centring-critical-top-align.md) |

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
