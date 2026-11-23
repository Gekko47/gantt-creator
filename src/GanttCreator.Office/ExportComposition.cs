using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Office;

/// <summary>
/// The export composition's width and height in points, recorded by a completed
/// refresh so the Ribbon preview can calculate pixel dimensions without a host
/// call (R8.7 D4, owner ruling 2026-11-23).
/// </summary>
/// <remarks>
/// <para>
/// <b>It describes the last completed render, not the table.</b> The refresh
/// records it after a successful reconciliation; a refused pass records nothing,
/// so the Ribbon keeps showing the aspect of the chart that is actually on screen.
/// The export command never trusts it — it resolves its own bounds from a fresh
/// preflight — so the cache is presentation-only and a mid-flight table change can
/// never corrupt an export.
/// </para>
/// <para>
/// <b>Validated at construction.</b> A non-finite or non-positive extent is a
/// caller error (the <see cref="RectD"/> guard style): the resolver can produce a
/// zero extent for a degenerate composition, and
/// <see cref="ExportComposition.TryResolveAspect"/> filters that to
/// <see langword="null"/> before constructing this value.
/// </para>
/// </remarks>
public readonly record struct ExportAspectPt
{
    /// <summary>Creates a validated export aspect.</summary>
    /// <param name="widthPt">The composition width in points. Must be finite and greater than zero.</param>
    /// <param name="heightPt">The composition height in points. Must be finite and greater than zero.</param>
    /// <exception cref="ArgumentOutOfRangeException">A dimension is not finite or not greater than zero.</exception>
    public ExportAspectPt(double widthPt, double heightPt)
    {
        if (!double.IsFinite(widthPt) || widthPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(widthPt),
                widthPt,
                "The export aspect width must be finite and greater than zero.");
        }

        if (!double.IsFinite(heightPt) || heightPt <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(heightPt),
                heightPt,
                "The export aspect height must be finite and greater than zero.");
        }

        WidthPt = widthPt;
        HeightPt = heightPt;
    }

    /// <summary>The composition width in points.</summary>
    public double WidthPt { get; }

    /// <summary>The composition height in points.</summary>
    public double HeightPt { get; }

    /// <summary>Whether a usable aspect can be built from these extents.</summary>
    /// <param name="widthPt">The candidate width in points.</param>
    /// <param name="heightPt">The candidate height in points.</param>
    /// <returns><see langword="true"/> when both extents are finite and greater than zero.</returns>
    /// <remarks>
    /// The resolver's bounds are finite and non-negative by <see cref="RectD"/>'s
    /// construction, so the reachable rejection here is the zero extent a
    /// degenerate composition resolves to.
    /// </remarks>
    public static bool IsUsable(double widthPt, double heightPt) =>
        double.IsFinite(widthPt)
        && widthPt > 0
        && double.IsFinite(heightPt)
        && heightPt > 0;
}

/// <summary>
/// Resolves the export composition's aspect host-free, or <see langword="null"/>
/// when the composition cannot be built (R8.7 D4).
/// </summary>
/// <param name="events">The validated events the export scene is composed from.</param>
/// <param name="settings">The workbook's effective settings map.</param>
/// <param name="registry">The workbook's style registry.</param>
/// <param name="grid">The measured panel grid.</param>
/// <param name="themedMetrics">The workbook's themed metric values.</param>
/// <returns>The resolved aspect, or <see langword="null"/> on any refusal.</returns>
/// <remarks>
/// The delegate exists so the orchestrator's contract test can drive the aspect
/// path deterministically without a host; production passes
/// <see cref="ExportComposition.TryResolveAspect"/>.
/// </remarks>
public delegate ExportAspectPt? ExportAspectResolve(
    IReadOnlyList<GanttEvent> events,
    IReadOnlyDictionary<string, string> settings,
    GanttStyleRegistry registry,
    PanelCellGrid grid,
    IReadOnlyDictionary<string, double> themedMetrics);

/// <summary>
/// The shared export-composition rule (R7.6/R8.7): which composition the export
/// covers, and the aspect that composition resolves to.
/// </summary>
/// <remarks>
/// <para>
/// <b>One authority for the selection.</b> The exact same rule
/// (<see cref="SelectExportSelection"/>) decides the export coverage for
/// Copy Editable, Send to PowerPoint, the preview's aspect and the Export PNG
/// command. It moved here from the AddIn pre-flight so the Office-side aspect
/// record and the AddIn-side commands cannot disagree about which composition
/// they describe.
/// </para>
/// <para>
/// <b>The aspect composition mirrors the preflight's steps.</b>
/// <see cref="TryResolveAspect"/> applies the same export-request factory, the
/// same selection rule and the same panel-theme resolution the pre-flight uses;
/// the equality for the same inputs is pinned by a test.
/// </para>
/// </remarks>
public static class ExportComposition
{
    /// <summary>
    /// The settings key that selects the export composition (catalogue default
    /// <c>TRUE</c>).
    /// </summary>
    public const string ExportIncludeDataPanelSettingKey = "ExportIncludeDataPanel";

    /// <summary>Maps the stored inclusion setting to the export selection.</summary>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <returns>Chart-only on an explicit case-insensitive <c>FALSE</c>; otherwise table-plus-chart.</returns>
    /// <remarks>
    /// Fail-open toward the catalogue default (<c>TRUE</c>), matching the
    /// automatic-refresh parse precedent: a typo cannot silently change what an
    /// export covers, and the catalogue remains the key's authority.
    /// </remarks>
    public static ExportSelection SelectExportSelection(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return settings.TryGetValue(ExportIncludeDataPanelSettingKey, out string? stored)
            && string.Equals(stored?.Trim(), "FALSE", StringComparison.OrdinalIgnoreCase)
                ? ExportSelection.ChartOnly
                : ExportSelection.TablePlusChart;
    }

    /// <summary>Resolves the export composition's aspect host-free.</summary>
    /// <param name="events">The validated events the export scene is composed from.</param>
    /// <param name="settings">The workbook's effective settings map.</param>
    /// <param name="registry">The workbook's style registry.</param>
    /// <param name="grid">The measured panel grid.</param>
    /// <param name="themedMetrics">The workbook's themed metric values.</param>
    /// <returns>The resolved aspect, or <see langword="null"/> when the composition is refused.</returns>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <remarks>
    /// Every expected refusal is a <see langword="null"/> rather than an
    /// exception: an empty table, a degenerate bounds and an unresolvable style
    /// are ordinary states of a workbook that has not rendered an export yet, and
    /// the Ribbon treats all of them as the not-available preview state.
    /// </remarks>
    public static ExportAspectPt? TryResolveAspect(
        IReadOnlyList<GanttEvent> events,
        IReadOnlyDictionary<string, string> settings,
        GanttStyleRegistry registry,
        PanelCellGrid grid,
        IReadOnlyDictionary<string, double> themedMetrics)
    {
        ArgumentNullException.ThrowIfNull(events);
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(registry);
        ArgumentNullException.ThrowIfNull(grid);
        ArgumentNullException.ThrowIfNull(themedMetrics);

        SceneBuildRequestOutcome request = EditableExportRequest.Create(
            events,
            settings,
            registry,
            grid,
            themedMetrics);
        if (!request.Succeeded || request.Request is not { } buildRequest)
        {
            return null;
        }

        ExportSelection selection = SelectExportSelection(settings);
        ExportBoundsOutcome resolution = ExportBoundsResolver.TryResolve(
            buildRequest,
            selection == ExportSelection.TablePlusChart
                ? ExportPanelTheme.Resolve(themedMetrics)
                : null,
            selection);
        if (!resolution.Succeeded || resolution.Bounds is not { } bounds)
        {
            return null;
        }

        RectD extent = bounds.Bounds;
        return ExportAspectPt.IsUsable(extent.Width, extent.Height)
            ? new ExportAspectPt(extent.Width, extent.Height)
            : null;
    }
}
