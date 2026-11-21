# ADR-0069 — Typed PowerPoint PIA and proxy ownership

- **Status:** Accepted
- **Date:** 2026-11-21
- **Work item:** `R7.1` ([`docs/work-items/R7.1-powerpoint-adapters.md`](../work-items/R7.1-powerpoint-adapters.md))
- **Supersedes:** the late-bound `dynamic` route used by the R6.7 test fixture
- **Related:** ADR-0004 (Excel-DNA pins), ADR-0020 (application-state scope), ADR-0066 (copy after delete)

## Context

Phase 7 places the composed Gantt on a PowerPoint slide. The owner directed statically typed
interop on 2026-11-21 and rejected the late-bound `dynamic` route: the R6.7 fixture's
`Microsoft.CSharp` binder calls are acceptable in a throwaway spike but not on a production
or test path. D-G6 therefore required a package decision **before** the first adapter file:
either `ExcelDna.Interop` already supplies the PowerPoint PIA, or a package is pinned and
recorded here.

`ExcelDna.Interop` 16.0.0 does **not** supply it. Its restored package file list is the Excel
PIA, `Microsoft.Vbe.Interop` and `office.dll` only ([`Directory.Packages.props`](../../Directory.Packages.props)).

`Microsoft.Office.Interop.PowerPoint` has **no 16.x release on nuget.org**. The published
versions are `12.0.4518.1014`, `15.0.4420.1017` and `15.0.4420.1018`. The highest version is a
repackaging of the Microsoft-signed Office 2013 PIA (assembly version `15.0.0.0`), and it ships
only `Microsoft.Office.Interop.PowerPoint.dll` — no `office.dll`, no dependencies.

## Decision

### D1 — Typed PIA, pinned at the highest published version

`Microsoft.Office.Interop.PowerPoint` **15.0.4420.1018** is pinned in
[`Directory.Packages.props`](../../Directory.Packages.props) and referenced by
`GanttCreator.Office` and `GanttCreator.Office.IntegrationTests`. The planned 16.0.0 pin is
impossible because the version does not exist; this is an availability fact, not a preference.
The reference host is Microsoft 365 MSO, version 2610, build 16.0.20527.20196, 64-bit, which
already ran the R6.7 `ppPasteShape` probe.

### D2 — Embedded interop types, ordered between the Excel-DNA targets

`ExcelDna.Interop.targets` marks `Microsoft.Office.Interop.Excel`, `Microsoft.Vbe.Interop` and
`office` with `EmbedInteropTypes=true`. The PowerPoint PIA's indirect reference to `office 15`
otherwise creates a real reference and fails with **CS1762** against the embedded one. The
Office project therefore sets `EmbedInteropTypes=true` for
`Microsoft.Office.Interop.PowerPoint` in a target ordered
`AfterTargets="LinkExcelDnaInteropEmbeddableAssemblies"` and
`BeforeTargets="FindReferenceAssembliesForReferences"`. One embedding policy covers every
interop assembly. The Release pack log confirms the PowerPoint PIA is additionally packed
into the XLL as a managed resource, so nothing has to be deployed beside it.

### D3 — Primitive ports, no late binding

`IPowerPointApplicationPort` exposes only primitives and records; no interop type appears in a
signature. Phase-7 production and test paths carry no `dynamic` and no `Microsoft.CSharp`
reference. Three structural guards keep this true: a source scan for both, a source scan for
`Activate`, and a reflective scan over every `IPowerPoint*` interface.

### D4 — Ownership: tracked, reverse order, exactly once

`PowerPointProxyOwnership` records each proxy the add-in acquires and releases it exactly once,
newest first, with `Marshal.ReleaseComObject`. It never calls
`Marshal.FinalReleaseComObject`, never force-releases a shared root, and never quits or closes a
pre-existing PowerPoint or presentation. A refused release is recorded once and does not escape.
The adapter keeps the application proxy in one root tracker and releases per-read proxies in a
transient tracker's `finally`.

### D5 — Attach-or-start is measured, never assumed

A typed `new Application()` attaches to a running instance or starts a stopped one. The adapter
captures the `POWERPNT` process set before and after the call and reports `StartedByUs` only when
the set was empty before and non-empty after. `Marshal.GetActiveObject` is not used: Microsoft
documents it for .NET Framework only, and the typed single-instance call covers both cases.

### D6 — Read, never activate; the surface cannot save

The active presentation, window, view and slide are read through `Application.ActiveWindow`,
`DocumentWindow.View` and `View.Slide`. No `Activate` is called on a slide, window or
presentation. No `Save`, `SaveAs`, `Close`, `Quit` or `QuitAll` member exists on any
`IPowerPoint*` interface, so R7.6's ownership proof is structural rather than a sequence hope.

## Host probe (2026-11-21, build 16.0.20527.20196, x64)

Measured by [`scripts/probe-powerpoint-host.ps1`](../../scripts/probe-powerpoint-host.ps1):

- PowerPoint was not running; the typed `new Application()` started it, `StartedByUs` was
  true, one `POWERPNT` process was created, and the probe left zero orphan processes.
- `Application.Version` is `16.0`, so the Office 2013-era PIA automates the current host.
- A windowless presentation has no `ActiveWindow`; its `Presentation.ReadOnly` is
  `msoFalse`; `Slides.Add` takes the slide count 0 to 1.
- `DisplayAlerts` reads 1 (`ppAlertsNone`) as this host's default.
- A missing member returns `0x80020006` (`DISP_E_MEMBERNOTFOUND`), which is an
  unsupported-build signal and never a transient retry.

The production-scale paste range and the three group-resize routes need the staged Excel
group, so they remain R7.3/R7.4 live-gate work.

## Consequences

**Positive.** The adapter is compiler-checked against the typed object model; a renamed or
mistyped member fails the build rather than a binder call at runtime. The release order and the
no-save rule are testable without a host. The AddIn stays interop-free.

**Negative and risks.**

1. The package is a community repackaging of Microsoft-signed assemblies (NuGet owner `e-bits`,
   described by the package itself as unsupported). The assemblies are Microsoft-signed, but the
   package is not published by Microsoft. This is a supply-chain risk the owner accepts by
   accepting this ADR; the version and hash are pinned and the lock files are committed.
2. The PIA is Office 2013-era (`15.0.0.0`) while `ExcelDna.Interop` supplies `16.0.0.0` of
   `office` and `Vbe.Interop`. .NET binds by simple name and accepts the higher assembly; the
   Release build and the guard tests prove compile-time compatibility. The live gate (R7.2) must
   prove runtime compatibility on the reference build.
3. Embedding the interop types means the PIA assembly is not deployed; a future need to change
   embedding policy is a build change, not a product change.

**Rejected alternatives.** Late-bound `dynamic` (owner direction);
`Marshal.GetActiveObject` (not available on .NET 10); an `oleaut32!GetActiveObject` P/Invoke
(the typed call is simpler and documented); a direct `<Reference>` to a machine's GAC copy
(not reproducible and not offline-deterministic); a third-party wrapper such as NetOffice (a
new abstraction for a single consumer).

## Rollback

Both the package and the adapter are additive. `git revert` of the R7.1 commits restores the
previous state; no worksheet, schema or catalogue value is involved.
