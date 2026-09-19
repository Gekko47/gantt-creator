# Repository map

Navigation only; [AGENTS.md](../AGENTS.md) retains requirement precedence. Verify current source before editing.
Source-reviewed 2026-09-18; this date does not detect subsequent or uncommitted changes.

| Project | Current responsibility and entry points |
| --- | --- |
| [AddIn](../src/GanttCreator.AddIn/) | Lifecycle: AddInHost; Ribbon: GanttRibbon / RibbonStateService; diagnostics: DiagnosticsService; errors: CommandBoundary; commands: InitialiseSheetCommand (sheet initialisation) |
| [Core](../src/GanttCreator.Core/) | PointD, VersionInfo, Logging, workbook schema contract (GanttTableColumns/GanttSchemaVersion, GanttWorkbookContract), the code-owned EntityTypeCatalog, and stable row identity (GanttRowId); no Office or UI dependency |
| [Office](../src/GanttCreator.Office/) | IExcelApplicationAdapter / ExcelApplicationAdapter: active-workbook facts and event subscriptions; IWorkbookInitialiser / ExcelWorkbookInitialiser: sheet initialisation (adopt/create, headers, table, config sheet, plot anchor); IGanttTableReader / ExcelGanttTableReader: bulk-read tblGanttData via Value2 into Core GanttRowDto (R2.4) |
| [Raster](../src/GanttCreator.Raster/) | ExportSize: width parsing and pixel dimensions, not PNG encoding |

A second table records the project-to-project references (the dependency graph); the first table records current responsibilities and entry points.

| Project | References |
| --- | --- |
| [AddIn](../src/GanttCreator.AddIn/) | Core, Office, Raster |
| [Core](../src/GanttCreator.Core/) | (none) |
| [Office](../src/GanttCreator.Office/) | Core |
| [Raster](../src/GanttCreator.Raster/) | Core |

[Tests](../tests/) mirror these areas; Architecture.Tests checks boundaries; Office.IntegrationTests requires live Office. R2.2 added `tests/GanttCreator.Office.ContractTests/WorkbookInitialiserTests.cs` (Moq PIA, no live Office); the live-Excel path is tagged `tests/GanttCreator.Office.IntegrationTests/InitialiseSheetIntegrationTests.cs` (`OfficeIntegration`). R2.4 added `src/GanttCreator.Core/GanttRowDto.cs` + `ExcelCellConverter.cs`, `src/GanttCreator.Office/IGanttTableReader.cs` + `ExcelGanttTableReader.cs` + `GanttTableReadOutcome.cs`; contract tests `tests/GanttCreator.Office.ContractTests/GanttTableReaderTests.cs`, live tests `tests/GanttCreator.Office.IntegrationTests/GanttTableReaderIntegrationTests.cs`, Core tests `tests/GanttCreator.Core.Tests/ExcelCellConverterTests.cs` + `GanttRowDtoTests.cs`.
Scene/layout, chart rendering, PowerPoint transfer, and PNG encoding remain planned: consult [architecture](02-ARCHITECTURE.md) and [roadmap](03-ROADMAP.md).
Read the current [work item](work-items/) and [STATUS](STATUS.md); use scoped live symbol queries, narrowing truncated results. A symbol outline is only a locator; treat it as navigation aid, not as compiler or reference evidence — verify any finding against current source.
From the repository root: `pwsh ./scripts/check-md-links.ps1` checks links, not semantic accuracy.
After staging intended changes: `pwsh ./scripts/verify-quick.ps1`; after committing with a clean tree: `pwsh ./scripts/verify.ps1`.
Live Office checks are separate: `pwsh ./scripts/verify-office.ps1` when required; see [verification policy](05-GIT-QUALITY.md).
Update this map in the same change as mapped responsibilities, entry points, test locations or commands; inspect working-tree edits before relying on it.
[Implementation plan](work-items/REPO-MAP-navigation.md). No persistent symbol cache or automatic freshness checker is installed.
