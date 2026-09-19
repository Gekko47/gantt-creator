// The Office contract tests substitute the adapter's internal virtual
// indexer-seam members (ExcelWorkbookInitialiser.GetSheetAt/GetTableAt/
// GetHeaderRange); expression trees cannot contain COM indexed properties, so
// the test assembly must see internals to derive the testable initialiser.
[assembly: System.Runtime.CompilerServices.InternalsVisibleTo("GanttCreator.Office.ContractTests")]
