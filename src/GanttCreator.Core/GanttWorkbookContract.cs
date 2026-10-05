namespace GanttCreator.Core;

/// <summary>
/// Workbook-level contract constants shared by the Office adapter that
/// creates the initialised state and the tests that pin it. These values are
/// part of the workbook schema: changing any of them is a schema migration
/// (see <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>).
/// </summary>
public static class GanttWorkbookContract
{
    /// <summary>
    /// The exact name of the single add-in-managed configuration worksheet.
    /// It must be the only helper sheet in the workbook and must be created
    /// with <c>xlSheetVeryHidden</c> visibility.
    /// </summary>
    public const string ConfigSheetName = "_GanttCreatorConfig";

    /// <summary>
    /// The deterministic label given to the visible Gantt worksheet on
    /// initialise — both when a blank active worksheet is adopted (renamed)
    /// and when a new worksheet is created. A name collision takes Excel's
    /// own <c>" (n)"</c> suffix convention.
    /// </summary>
    public const string GanttSheetLabel = "Gantt Data";

    /// <summary>
    /// The sheet-scoped defined name recording the plot anchor cell. Written
    /// at initialise time; the renderer derives the anchor from the live
    /// table and treats stored-vs-derived disagreement as an integrity
    /// finding. Sheet-scoped so a future multi-Gantt-sheet extension needs no
    /// name migration.
    /// </summary>
    public const string PlotAnchorDefinedName = "GanttCreator.PlotAnchor";

    /// <summary>
    /// The workbook-scoped defined name containing the exact Type display-name
    /// cells materialised in <c>tblGanttTypes</c>.
    /// </summary>
    public const string TypeOptionsDefinedName = "GanttCreator.TypeOptions";

    /// <summary>
    /// The title written into the reserved row when the settings catalogue cannot
    /// supply one (ADR-0030 D6).
    /// </summary>
    /// <remarks>
    /// <b>This is a fallback, not an authority.</b> The live value comes from the
    /// <c>ChartTitle</c> setting via <c>GanttCatalogues.SettingDefault</c>, which is
    /// the same source the settings table and the export title entity use. This
    /// constant exists only so Initialise can write a sensible title into a workbook
    /// whose settings table has not been read yet; naming it
    /// <c>Gantt Chart</c> matches entity guide §2's documented nonblank default so the
    /// fallback and the catalogue agree today.
    /// </remarks>
    public const string DefaultChartTitle = "Gantt Chart";
}
