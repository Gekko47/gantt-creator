# ADR-0004 — Pin Excel-DNA package versions to AddIn 1.9.0, Integration 1.9.0, Interop 16.0.0

- **Status**: Accepted
- **Date**: 2026-09-04
- **Context**: `docs/01-ENVIRONMENT.md` declares "Excel-DNA 1.9.0" and the first `Directory.Packages.props` draft pinned `ExcelDna.AddIn 1.9.0` and `ExcelDna.Interop 1.9.0`. Restore failed with `NU1603` because `ExcelDna.Interop 1.9.0` is not a published package version. Verified directly on nuget.org: the `ExcelDna.Interop` package version sequence is `14.0.1`, `15.0.0`, `15.0.1`, `16.0.0`; the latest stable is `16.0.0`, which targets `net6.0-windows7.0` and is declared compatible with `net10.0-windows`. The Excel-DNA project itself is on version 1.9.0 across `ExcelDna.AddIn` and `ExcelDna.Integration` (the latter is a transitive of `AddIn` but pinned here so a future NuGet resolution cannot surprise us).
- **Decision**: Pin three packages in `Directory.Packages.props`:
  - `ExcelDna.AddIn` 1.9.0 (latest stable, produces the `.xll`)
  - `ExcelDna.Integration` 1.9.0 (pinned explicitly; not a direct project reference)
  - `ExcelDna.Interop` 16.0.0 (latest stable; provides the Microsoft Office PIAs)
  Any future bump requires an ADR.
- **Consequences**: The .NET 10 SDK is happy (the interop package is declared compatible with `net10.0-windows`). `ExcelDna.Integration` is consumed transitively and the pin is documentation, not a behaviour change. The Interop version is *not* the same as the AddIn version because the Interop and AddIn packages follow separate SemVer; this is a permanent gotcha worth flagging in the README.
- **Alternatives considered**: Pinning both to `1.9.0` (rejected — `ExcelDna.Interop 1.9.0` does not exist). Using the legacy `Excel-DNA.Interop` package (rejected — the nuget.org page explicitly deprecates it in favour of `ExcelDna.Interop`). Bumping the entire kit to a future `ExcelDna.AddIn 1.10.x` preview (rejected — the preview is from July 2026, and we are on 1.9.0 stable per the environment skill).
