# ADR-0002 — Commit `.clinerules/*` and `AGENTS.md` as canonical sources

- **Status**: Accepted
- **Date**: 2026-09-04
- **Context**: The prior `.gitignore` ignored `.clinerules/*` and `AGENTS.md`, treating them as agent-local state. The current working tree contains seven `.clinerules/0N-*.md` files and a root `AGENTS.md`. The roadmap and the agent contract both treat these files as the always-on operating contract, not as ephemeral state.
- **Decision**: Commit `.clinerules/*` and `AGENTS.md`. Keep `.cline/*` ignored as the agent scratch area. The seven rules are mirrored into `docs/0N-*.md` so the in-repo source of truth is human-reviewable, while `.clinerules/` is the Cline-loader view.
- **Consequences**: Any change to the agent contract is now visible in `git log` and reviewable. Local divergence between the rule files and `docs/` is a defect, caught by the documentation-validation gate in R0.6. The cost is that future contributors must remember to update both copies when the contract changes.
- **Alternatives considered**: Keep both ignored and copy the content into the README (rejected — chat-style drift). Keep only `docs/` and remove `.clinerules/` (rejected — the Cline loader expects the always-on view at `.clinerules/`).
