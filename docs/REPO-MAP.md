# Repository map

Navigation only; [AGENTS.md](../AGENTS.md) retains requirement precedence. Verify current source before editing.
Source-reviewed 2026-09-17; this date does not detect subsequent or uncommitted changes.

| Project | Current responsibility and entry points |
| --- | --- |
| [AddIn](../src/GanttCreator.AddIn/) | Lifecycle: AddInHost; Ribbon: GanttRibbon / RibbonStateService; diagnostics: DiagnosticsService; errors: CommandBoundary |
| [Core](../src/GanttCreator.Core/) | PointD, VersionInfo, Logging; no Office or UI dependency |
| [Office](../src/GanttCreator.Office/) | IExcelApplicationAdapter / ExcelApplicationAdapter: active-workbook facts and event subscriptions |
| [Raster](../src/GanttCreator.Raster/) | ExportSize: width parsing and pixel dimensions, not PNG encoding |

[Tests](../tests/) mirror these areas; Architecture.Tests checks boundaries; Office.IntegrationTests requires live Office.
Scene/layout, chart rendering, PowerPoint transfer, and PNG encoding remain planned: consult [architecture](02-ARCHITECTURE.md) and [roadmap](03-ROADMAP.md).
Read the current [work item](work-items/) and [STATUS](STATUS.md); use scoped live symbol queries, narrowing truncated results.
From the repository root: `pwsh ./scripts/check-md-links.ps1` checks links, not semantic accuracy.
After staging intended changes: `pwsh ./scripts/verify-quick.ps1`; after committing with a clean tree: `pwsh ./scripts/verify.ps1`.
Live Office checks are separate: `pwsh ./scripts/verify-office.ps1` when required; see [verification policy](05-GIT-QUALITY.md).
Update this map in the same change as mapped responsibilities, entry points, test locations or commands; inspect working-tree edits before relying on it.
[Implementation plan](work-items/REPO-MAP-navigation.md). No persistent symbol cache or automatic freshness checker is installed.
