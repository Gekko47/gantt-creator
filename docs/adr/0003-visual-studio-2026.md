# ADR-0003 — Visual Studio 2026 is the verification IDE; the 2022 baseline is relaxed

- **Status**: Accepted
- **Date**: 2026-09-04
- **Context**: The environment skill (`docs/01-ENVIRONMENT.md`) lists Visual Studio 2022 17.14+ as the supported IDE. The actual installed IDE is Visual Studio 2026 (18) Community at `C:\Program Files\Microsoft Visual Studio\18\Community`. The Excel-DNA 1.9 first-run recipe, the F5 external-program launch, and the Copilot Agent Mode integration are all expected to work on VS 2026 with the same shape.
- **Decision**: Use Visual Studio 2026 (18) Community as the verification IDE. The R1.1 Office gate must record the actual VS build number in the PR evidence block. The "VS 2022 17.14+" baseline is recorded as a portability target, not a hard requirement, until the first release candidate is produced on a VS 2022 machine.
- **Consequences**: First Office gate evidence is reproducible on this machine. Future release candidates must be rehearsed on the original VS 2022 17.14+ baseline to prove portability. Tracked as L2 in `docs/KNOWN-LIMITATIONS.md`.
- **Alternatives considered**: Downgrade the IDE to VS 2022 (rejected — the user has installed VS 2026 and the difference is irrelevant to the Excel-DNA host contract). Require the CI runner to use VS 2022 (rejected — out of scope until the R10.5 evidence manifest).
