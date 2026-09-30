using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>Why a panel-grid measurement refused to return a grid.</summary>
public enum PanelGridRefusalReason
{
    /// <summary>The application object or active workbook was absent.</summary>
    NoActiveWorkbook = 0,

    /// <summary>The Gantt worksheet or its <c>tblGanttData</c> table was missing.</summary>
    TableMissing = 1,

    /// <summary>The target worksheet is protected.</summary>
    TargetProtected = 2,

    /// <summary>
    /// The host reported a non-positive or non-finite column width or row height,
    /// which <c>PanelCellGrid.TryCreate</c> would refuse anyway. The adapter
    /// reports it here rather than letting the Core refusal surface untyped.
    /// </summary>
    InvalidMeasurement = 3,
}

/// <summary>The typed result of one panel-grid measurement.</summary>
/// <param name="Grid">The measured grid on success, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record PanelGridOutcome(PanelCellGrid? Grid, PanelGridRefusalReason? Refusal)
{
    /// <summary>Gets whether a grid was measured.</summary>
    public bool Succeeded => Grid is not null;

    /// <summary>Creates a successful outcome.</summary>
    /// <param name="grid">The measured grid.</param>
    /// <returns>The successful outcome.</returns>
    public static PanelGridOutcome Ok(PanelCellGrid grid) => new(grid, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static PanelGridOutcome Refused(PanelGridRefusalReason refusal) => new(null, refusal);
}

/// <summary>
/// Narrow, read-only port over the live measurement of the data panel's cell
/// grid: the exact point widths of the included <c>tblGanttData</c> columns, the
/// height of <em>each</em> body row, and the header row's own height.
/// </summary>
/// <remarks>
/// <para>
/// The surface is deliberately primitive — no interop type crosses the port —
/// so the AddIn compilation never names <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// <para>
/// Row heights are returned <b>per row, in body order</b>, not as one aggregate.
/// A worksheet body does not have to be uniform, and entity guide §3 requires the
/// panel to reproduce the exact measured cell bounds, so a single height could only
/// refuse a legitimate table or lay it out wrong. The header row is measured
/// separately because §4 makes it follow the live header-cell bounds.
/// </para>
/// <para>
/// This port <strong>measures only</strong>. R3.11's <c>PanelBuilder</c> and
/// <c>SceneBuilder</c> consume the result and never re-measure or default a
/// missing measurement, so this is the single source of live panel geometry.
/// A caller may not substitute a guessed width or height.
/// </para>
/// <para>
/// The port is read-only: it never writes cells, changes column widths, or
/// alters application state. It therefore does <b>not</b> refuse a protected
/// target — worksheet protection blocks writes, not reads — and the caller
/// enforces protection at the Refresh/write boundary.
/// </para>
/// </remarks>
public interface IPanelGridMeasurementPort
{
    /// <summary>Measures the live data panel's cell grid.</summary>
    /// <param name="includedColumns">
    /// The schema display names of the columns to measure, in the order the
    /// panel must use. Order is significant — the emitted cell bounds depend on
    /// it — so it is preserved rather than sorted.
    /// </param>
    /// <returns>
    /// The measured grid, or the refusal reason. On a refusal nothing was
    /// mutated.
    /// </returns>
    PanelGridOutcome Measure(IReadOnlyList<string> includedColumns);
}
