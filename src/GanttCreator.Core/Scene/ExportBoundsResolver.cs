namespace GanttCreator.Core.Scene;

/// <summary>The typed result of resolving export bounds.</summary>
/// <param name="Bounds">The resolved bounds, or <see langword="null"/>.</param>
/// <param name="Refusal">The export refusal, or <see langword="null"/>.</param>
/// <param name="SceneRefusal">
/// The scene builder's own refusal when <paramref name="Refusal"/> is
/// <see cref="ExportBoundsRefusal.SceneBuildRefused"/>, so the underlying cause is not
/// lost behind the export classification.
/// </param>
public sealed record ExportBoundsOutcome(
    ExportBounds? Bounds,
    ExportBoundsRefusal? Refusal,
    SceneBuilderRefusal? SceneRefusal)
{
    /// <summary>Gets whether the resolution succeeded.</summary>
    public bool Succeeded => Bounds is not null;
}

/// <summary>
/// Resolves the bounds a Copy Editable composition covers for each export selection,
/// creating no shapes.
/// </summary>
/// <remarks>
/// <para>
/// <b>One composition authority.</b> The bounds are the chart bounds
/// <see cref="SceneBuilder"/> derives from the frame/band builder's union — not a second
/// union computed here. That is what keeps this model and R6.3's staged composition from
/// disagreeing: both read the scene the same orchestrator builds, so the grouped-export
/// bounds check and Phase 8's crop consume this value rather than recomputing it.
/// </para>
/// <para>
/// <b>The selection rides the existing seam.</b> The model pins the composition profile
/// to <see cref="SceneCompositionProfile.EditableExport"/> and supplies
/// <see cref="SceneBuildRequest.Panel"/> only for
/// <see cref="ExportSelection.TablePlusChart"/>. The profile is part of the model's
/// contract rather than the caller's choice: a live profile omits the drawn title band
/// (ADR-0030 D6) and would therefore produce different bounds for the same chart.
/// </para>
/// <para>
/// <b>Refusals, not empties.</b> A table-plus-chart composition with no visible row is
/// refused rather than exported as the chart alone: entity guide section 1 keeps an
/// empty data table invalid <em>as an export</em>, even though the live path draws a
/// calendar-year chart for one (ADR-0050, R5.18j). The two rules serve different
/// consumers, and this refusal is export-only by contract. A chart-only composition is
/// not subject to the table rule; if its chart is empty, the scene builder's own
/// <see cref="SceneBuilderRefusal.EmptyEvents"/> is surfaced instead.
/// </para>
/// <para>
/// <b>Zero visible columns degrade, they do not refuse (ADR-0065 D4).</b> When every
/// user column is hidden the measured grid sums to zero width. The panel is then not
/// supplied at all, so the table-plus-chart export is exactly the chart-only export:
/// the panel contributes no bounds, no phantom-width rectangle is composed, and the
/// frame builder does not see a zero-width panel it would refuse as invalid geometry.
/// </para>
/// </remarks>
public static class ExportBoundsResolver
{
    /// <summary>Resolves the export bounds for one selection.</summary>
    /// <param name="request">The composition request the export scene would be built from.</param>
    /// <param name="panelTheme">
    /// The resolved data-panel styles. Required by, and only used for,
    /// <see cref="ExportSelection.TablePlusChart"/>.
    /// </param>
    /// <param name="selection">The export selection the user chose.</param>
    /// <returns>The zero-origin bounds with their placement delta, or a typed refusal.</returns>
    public static ExportBoundsOutcome TryResolve(
        SceneBuildRequest? request,
        PanelTheme? panelTheme,
        ExportSelection selection
    )
    {
        if (request is null)
        {
            return Refused(ExportBoundsRefusal.NullRequest);
        }

        // Fail closed on an unrecognised value, for the same reason
        // SceneCompositionProfiles.DrawsDataPanel does: a member added to the enum
        // without a rule must not silently receive one mode's behaviour.
        if (!Enum.IsDefined(selection))
        {
            return Refused(ExportBoundsRefusal.InvalidSelection);
        }

        bool tablePlusChart = selection == ExportSelection.TablePlusChart;

        if (tablePlusChart && panelTheme is null)
        {
            return Refused(ExportBoundsRefusal.PanelThemeMissing);
        }

        // ADR-0065 D4: a table whose every user column is hidden is a supported
        // workbook state. A zero-width panel contributes nothing to the union, and the
        // frame builder refuses a zero-width panel rectangle as invalid geometry - so
        // the model expresses the degradation here, where it chooses whether a panel
        // is supplied at all, rather than composing a panel that cannot contribute.
        bool panelContributes = tablePlusChart
            && request.Grid is { } grid
            && grid.TotalWidthPt > 0;

        // "Empty" is a question about the visible table, not about the event list.
        // Section 3 reproduces every row in the panel, so a hidden row is still a
        // panel row - but a table with no visible row has nothing to export beside
        // the chart, and section 1 keeps that state invalid as an export.
        bool hasVisibleRow = request.Events is not null
            && request.Events.Any(@event => @event is not null && @event.Visible);

        if (tablePlusChart && !hasVisibleRow)
        {
            return Refused(ExportBoundsRefusal.EmptyVisibleTable);
        }

        // Invalid bounds refuse BEFORE any composition (work-item D2). A zero-extent
        // plot has no export rectangle to derive, and building first would report the
        // frame builder's refusal instead of the caller's bad input.
        if (request.PlotBounds is { } plotBounds
            && (plotBounds.Width <= 0 || plotBounds.Height <= 0))
        {
            return Refused(ExportBoundsRefusal.InvalidBounds);
        }

        // D4: the selection chooses whether the panel is supplied. The profile is the
        // export contract itself, so it is pinned rather than inherited from a caller
        // that may still be carrying a live template.
        SceneBuildRequest exportRequest = request with
        {
            Profile = SceneCompositionProfile.EditableExport,
            Panel = panelContributes ? panelTheme : null,
        };

        SceneBuildOutcome build = SceneBuilder.TryBuild(exportRequest);
        return build.Result is { } result
            ? new ExportBoundsOutcome(new ExportBounds(result.Scene.ChartBounds), null, null)
            : new ExportBoundsOutcome(null, ExportBoundsRefusal.SceneBuildRefused, build.Refusal);
    }

    private static ExportBoundsOutcome Refused(ExportBoundsRefusal refusal) => new(null, refusal, null);
}
