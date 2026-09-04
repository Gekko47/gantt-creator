# ADR-0000 — Adopt the Gantt Creator development kit as the source of truth

- **Status**: Accepted
- **Date**: 2026-09-04
- **Context**: The repository started as a legacy `Gantt Generator - v5.1.1.xlsm` macro workbook. A development kit was produced (`.cline/skills/*`, `.cline/roadmap/*`, `.START-HERE.md`) and the team needs to commit to the kit as the durable source of truth or continue reasoning from the macro.
- **Decision**: Adopt the kit. Mirror the seven rules into `docs/0N-*.md` so they are committed and human-reviewable. Keep `.clinerules/*` as the always-on Cline view. Keep `.cline/*` as the agent scratch area, ignored by Git. The `Gantt Generator - v5.1.1.xlsm` workbook is treated as a behaviour and visual reference, not as the architecture.
- **Consequences**: Every commit must satisfy the kit's quality gates (`scripts/verify-quick.ps1`, `scripts/verify.ps1`). The macro's behaviour is the *visual* contract, but the architecture is the kit's. Drift from the kit requires an ADR.
- **Alternatives considered**: Continue maintaining the macro (rejected — the user is paying for a redesign). Adopt a different architectural baseline (rejected — no approved alternative exists).
