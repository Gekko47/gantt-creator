// The Office contract tests substitute the adapter's internal virtual
// indexer-seam members (ExcelWorkbookInitialiser.GetSheetAt/GetTableAt/
// GetHeaderRange); expression trees cannot contain COM indexed properties, so
// the test assembly must see internals to derive the testable initialiser.
// The live Office integration test accesses the internal ValidationReportComposer
// to read the note sentinel prefix used in live assertions.
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("GanttCreator.Office.ContractTests")]
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("GanttCreator.Office.IntegrationTests")]
