# ADR-0009 — Named-style capability columns

- **Status**: Accepted
- **Date**: 2026-09-23
- **Context**: `Custom Activity` requires a named `StyleKey`, and R2.9 must validate its label and colour overrides. ADR-0007 fixed `tblGanttStyles` without a default label position, allowed label positions, or colour-override capability, so the required semantics are not representable and a renderer would have to invent them.
- **Decision**: Append `DefaultLabelPosition`, `AllowedLabelPositions`, and `ColourCapability` to `tblGanttStyles`. Built-in style rows project these values from their owning `EntityTypeCatalog` entry. User style rows carry and validate the same explicit fields. `Custom Activity` resolves capabilities only from the selected named style. Schema version remains 1 because configuration v1 was not shipped before R2.7 materialised it; older development workbooks are pre-release drift and must be recreated or explicitly repaired.
- **Consequences**: R2.9 can validate Custom Activity without a silent fallback, and R3 can consume one resolved style model. The writer/reader/hash/schema tests cover three additional fields. Later style creation must populate them rather than relying on renderer defaults.
- **Alternatives considered**: Hard-code span labels for Custom Activity (rejected: contradicts style-defined semantics); infer capability from non-empty colour cells (rejected: ambiguous and cannot represent label policy); defer Custom Activity support (rejected: it is an approved selectable Type and an explicit R2.9 obligation).
