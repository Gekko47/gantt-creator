# ADR-0013 — Chart-level scene ownership

- **Status**: Accepted
- **Date**: 2026-09-25
- **Context**: R3.2 scene primitives require a `GanttRowId` owner, but chart frame, background, time bands, grid lines, year/period headers, and title entities belong to the workbook rather than any visible table row. Reusing a fabricated row ID risks collision with real IDs and weakens the stable-identity contract; a nullable owner loses the explicit ownership discriminator required by later renderers and scene snapshots.
- **Decision**: Introduce an immutable `SceneOwnerId` with two closed kinds: a real row owner carrying the exact `GanttRowId`, or one deterministic reserved chart owner. Scene primitives and warnings use `SceneOwnerId`; row primitive IDs remain unchanged and chart primitive IDs begin with the reserved chart identity. Strict scene snapshots serialise the owner kind and value and reject missing, unknown, or malformed combinations.
- **Consequences**: Row ownership remains lossless and collision-safe. Workbook-level entities have explicit ownership and deterministic identity across Excel, editable export, and PNG renderers. The scene snapshot format changes before a committed golden scene baseline exists, so R3.2 tests and consumers must migrate together.
- **Alternatives considered**: Nullable owner ID (rejected because absence becomes ambiguous); fixed synthetic `GanttRowId` (rejected because it can collide with a valid row and pretends chart entities are rows); role strings without an owner model (rejected because ownership and primitive identity would diverge).
