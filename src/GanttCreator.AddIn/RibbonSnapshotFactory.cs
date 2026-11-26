using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>
/// The live inputs one snapshot capture reads. Every member is read at capture
/// time, so a port injected later is picked up by the next capture.
/// </summary>
/// <param name="ApplicationAdapter">Supplies the active-workbook fact, or null when unwired.</param>
/// <param name="LogAvailableSource">Supplies the log-availability fact, or null when unwired.</param>
/// <param name="SelectionIdentifier">Resolves the current selection, or null when unwired.</param>
/// <param name="TableReader">Reads the live table for the label contexts, or null when unwired.</param>
/// <param name="CatalogueReader">Reads the style registry, or null when unwired.</param>
/// <param name="LastTableRows">The rows the last successful read returned.</param>
/// <param name="CommandBoundary">The sink a kept-previous record is written through.</param>
internal sealed record SnapshotSources(
    IExcelApplicationAdapter? ApplicationAdapter,
    Func<bool?>? LogAvailableSource,
    ISelectionIdentifier? SelectionIdentifier,
    IGanttTableReader? TableReader,
    IConfigCatalogueReader? CatalogueReader,
    IReadOnlyList<GanttRowDto> LastTableRows,
    CommandBoundary CommandBoundary);

/// <summary>
/// Builds one <see cref="RibbonState"/> snapshot from the published snapshot and
/// the freshly probed facts.
/// </summary>
/// <remarks>
/// Every fact is sticky: a null probe keeps the published value, and a source that
/// throws keeps the entire previous snapshot. That policy is why a transient COM
/// failure cannot grey the surface, and it is preserved exactly here.
/// </remarks>
internal static class RibbonSnapshotFactory
{
    /// <summary>
    /// Captures the next snapshot. A source failure keeps
    /// <paramref name="previous"/> and records one technical line.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <param name="sheetInitialised">Whether the Gantt table is readable, or null when not determinable.</param>
    /// <param name="configurationPresent">Whether the configuration worksheet exists, or null when not determinable.</param>
    /// <param name="ganttSheetPresent">Whether a standard-labelled Gantt sheet exists, or null when not determinable.</param>
    /// <param name="criticalIntervalParentAvailable">Whether a row can own a Critical Interval, or null when not determinable.</param>
    /// <param name="activeSheetIsGanttSheet">Whether the active sheet carries the table, or null when not determinable.</param>
    /// <param name="hasGroups">Whether any row names a parent, or null when not determinable.</param>
    /// <returns>The snapshot to publish.</returns>
    internal static RibbonState Capture(
        SnapshotSources sources,
        RibbonState previous,
        bool? sheetInitialised,
        bool? configurationPresent = null,
        bool? ganttSheetPresent = null,
        bool? criticalIntervalParentAvailable = null,
        bool? activeSheetIsGanttSheet = null,
        bool? hasGroups = null)
    {
        ArgumentNullException.ThrowIfNull(sources);
        ArgumentNullException.ThrowIfNull(previous);

        // CA1031: the capture runs inside Ribbon callbacks and Excel events; a
        // source failure keeps the last known snapshot instead of propagating
        // into Excel.
#pragma warning disable CA1031
        try
        {
            bool? hasActiveWorkbook = sources.ApplicationAdapter?.HasActiveWorkbook() ?? null;
            bool? logAvailable = sources.LogAvailableSource?.Invoke() ?? null;

            // R5.16: resolve the mode from the three facts before the snapshot
            // is built. Each fact falls back to the published value so a
            // transient failure cannot flip the visible surface.
            bool workbookPresent = hasActiveWorkbook ?? previous.HasActiveWorkbook;
            bool initialised = sheetInitialised ?? previous.SheetInitialised;
            bool configPresent = configurationPresent ?? previous.ConfigurationPresent;
            bool ganttPresent = ganttSheetPresent ?? previous.GanttSheetPresent;
            bool criticalAvailable =
                criticalIntervalParentAvailable ?? previous.CriticalIntervalParentAvailable;

            // R5.16b: both new facts are sticky like every other captured fact:
            // an unreadable table keeps the published value instead of greying
            // or un-greying the whole surface mid-session.
            bool activeSheetIsGantt =
                activeSheetIsGanttSheet ?? previous.ActiveSheetIsGanttSheet;
            bool groupsPresent = hasGroups ?? previous.HasGroups;
            RibbonMode mode = ResolveMode(
                workbookPresent,
                initialised,
                configPresent,
                ganttPresent);

            // R5.5 ruling 6: the description-label dropdown is the first selection-gated
            // control, so its context is captured here rather than left at its
            // initial value. The capture is sticky like every other fact: a
            // transient failure keeps the previous context rather than greying
            // the dropdown mid-selection.
            DescriptionLabelContext descriptionLabelContext = CaptureDescriptionLabelContext(sources, previous);
            DateLabelContext dateLabelContext = CaptureDateLabelContext(sources, previous);
            DateLabelPlacementContext dateLabelPlacementContext = CaptureDateLabelPlacementContext(sources, previous);
            CustomStyleContext customStyleContext = CaptureCustomStyleContext(sources, previous);
            RowActionContext rowActionContext = CaptureRowActionContext(sources, previous);

            return DebugForceRibbonState.Apply(new RibbonState(
                HasActiveWorkbook: workbookPresent,
                LogAvailable: logAvailable ?? previous.LogAvailable,
                SheetInitialised: initialised,
                PlotStartAuto: previous.PlotStartAuto,
                PlotFinishAuto: previous.PlotFinishAuto,
                PlotStartDate: previous.PlotStartDate,
                PlotFinishDate: previous.PlotFinishDate,
                EffectivePlotStartDate: previous.EffectivePlotStartDate,
                EffectivePlotFinishDate: previous.EffectivePlotFinishDate,
                TimeScale: previous.TimeScale,
                Margin: previous.Margin,
                MarginCm: previous.MarginCm,
                Preset: previous.Preset,
                PlotWidthPt: previous.PlotWidthPt,
                PlotHeightPt: previous.PlotHeightPt,
                PeriodLabelFormat: previous.PeriodLabelFormat,
                DescriptionLabelSelectionAvailable: descriptionLabelContext.Available,
                SelectedDescriptionLabelPosition: descriptionLabelContext.Selected,
                DescriptionLabelOptions: descriptionLabelContext.Options,
                DateLabelSelectionAvailable: dateLabelContext.Available,
                SelectedDateLabelPosition: dateLabelContext.Selected,
                DateLabelOptions: dateLabelContext.Options,
                DateLabelPlacementSelectionAvailable: dateLabelPlacementContext.Available,
                SelectedDateLabelPlacement: dateLabelPlacementContext.Selected,
                DateLabelPlacementOptions: dateLabelPlacementContext.Options,
                CustomStyleSelectionAvailable: customStyleContext.Available,
                SelectedCustomStyleKey: customStyleContext.SelectedKey,
                CustomStyleOptions: customStyleContext.Options,
                ConfigurationPresent: configPresent,
                GanttSheetPresent: ganttPresent,
                Mode: mode,
                CriticalIntervalParentAvailable: criticalAvailable,
                RowActions: rowActionContext.Actions,
                ActiveCellInBodyRow: rowActionContext.ActiveCellInBodyRow,
                HasGroups: groupsPresent,
                ActiveSheetIsGanttSheet: activeSheetIsGantt));
        }
        catch (Exception exception)
        {
            // The sticky-snapshot policy is deliberate; the silence was not.
            // One record is the only evidence a future snapshot defect leaves.
            sources.CommandBoundary.WriteTechnicalRecord(
                $"gantt.ribbon snapshot kept-previous type={exception.GetType().FullName}");
            return previous;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Captures the structural row-action facts for the row under the active
    /// cell (R5.9b). Never throws: an unavailable selection or rule failure keeps
    /// the previous context.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private static RowActionContext CaptureRowActionContext(SnapshotSources sources, RibbonState previous)
    {
        RowActionContext unchanged = new(previous.RowActions, previous.ActiveCellInBodyRow);
        ISelectionIdentifier? identifier = sources.SelectionIdentifier;
        if (identifier is null)
        {
            return unchanged;
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.HasSelection)
            {
                return unchanged;
            }

            int? activeCellBodyIndex = facts.ActiveCellBodyIndex;
            return new RowActionContext(
                RowActionAvailability.Evaluate(sources.LastTableRows, activeCellBodyIndex),
                activeCellBodyIndex.HasValue);
        }
        catch
        {
            return unchanged;
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// Resolves the Ribbon mode from the three workbook facts (R5.16). A workbook
    /// that carries configuration but has lost its table is
    /// <see cref="RibbonMode.NeedsRepair"/>, not a first run.
    /// </summary>
    /// <param name="hasActiveWorkbook">Whether Excel has an active workbook.</param>
    /// <param name="sheetInitialised">Whether the Gantt table is readable.</param>
    /// <param name="configurationPresent">Whether <c>_GanttCreatorConfig</c> exists.</param>
    /// <param name="ganttSheetPresent">Whether a standard-labelled Gantt worksheet exists.</param>
    /// <returns>The resolved mode.</returns>
    private static RibbonMode ResolveMode(
        bool hasActiveWorkbook,
        bool sheetInitialised,
        bool configurationPresent,
        bool ganttSheetPresent) =>
        !hasActiveWorkbook
            ? RibbonMode.FirstRun
            : sheetInitialised
                ? configurationPresent
                    ? RibbonMode.Ready
                    : RibbonMode.NeedsRepair
                : configurationPresent || ganttSheetPresent
                    ? RibbonMode.NeedsRepair
                    : RibbonMode.FirstRun;

    /// <summary>
    /// Whether any row in the table can own a Critical Interval (R5.16).
    /// </summary>
    /// <param name="rows">The rows the table read returned.</param>
    /// <returns><see langword="true"/> when at least one row is an eligible parent.</returns>
    internal static bool HasEligibleParent(IReadOnlyList<GanttRowDto> rows)
    {
        if (rows.Count == 0)
        {
            return false;
        }

        Dictionary<string, int> childCounts = new(StringComparer.Ordinal);
        foreach (GanttRowDto row in rows)
        {
            string? parent = row.ParentId?.Trim();
            if (!string.IsNullOrEmpty(parent))
            {
                childCounts[parent] = childCounts.TryGetValue(parent, out int count) ? count + 1 : 1;
            }
        }

        foreach (GanttRowDto row in rows)
        {
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
            {
                continue;
            }

            GanttRowId? parentId = null;
            string? parentText = row.ParentId?.Trim();
            if (!string.IsNullOrEmpty(parentText)
                && (!GanttRowId.TryParse(parentText, out parentId) || parentId is null))
            {
                // A parent id that cannot be parsed makes the row's own depth
                // unknown, so the candidate is skipped rather than assumed
                // top-level: the insert would refuse it anyway.
                continue;
            }

            int childCount = 0;
            string? idText = row.Id?.Trim();
            if (!string.IsNullOrEmpty(idText))
            {
                _ = childCounts.TryGetValue(idText, out childCount);
            }

            if (AddChildPlanner.IsEligibleParent(type, parentId, childCount))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Captures the description-label dropdown's context from the selection
    /// identifier, the table reader, and the catalogue reader (R5.5 ruling 6).
    /// Never throws: an indeterminable selection keeps the previous context.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private static DescriptionLabelContext CaptureDescriptionLabelContext(
        SnapshotSources sources,
        RibbonState previous)
    {
        ISelectionIdentifier? identifier = sources.SelectionIdentifier;
        IGanttTableReader? reader = sources.TableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDescriptionLabelContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            // Custom Activity / Custom Milestone resolve their set from the
            // row's style (R2.7b/c; R5.6a: the catalogue's RequiresStyleKey
            // flag is the single authority); every other type reads its
            // catalogue definition. A failed catalogue read yields an empty
            // registry, which the projection turns into an empty option set:
            // the dropdown then offers only Default, and the command refuses
            // with StyleUnknown at execution.
            GanttStyleRegistry styles = GanttStyleRegistry.Empty;
            if (EntityTypeCatalog.GetDefinition(type) is { RequiresStyleKey: true }
                && sources.CatalogueReader is not null)
            {
                ConfigReadOutcome catalogues = sources.CatalogueReader.Read();
                if (catalogues.Succeeded)
                {
                    styles = catalogues.Styles;
                }
            }

            IReadOnlyList<GanttLabelPosition> options =
                EntityTypeCatalog.GetAllowedLabelPositions(type, row.StyleKey, styles);
            if (options.Count == 0)
            {
                return UnavailableDescriptionLabelContext(previous);
            }

            // The cell stores the position by name; blank resolves to the
            // style default / Auto cascade, so it reads as null.
            GanttLabelPosition? selected = EntityTypeCatalog.TryParseLabelSelection(
                row.LabelPositionText,
                out GanttLabelPosition? parsed)
                ? parsed
                : null;

            return new DescriptionLabelContext(true, selected, options);
        }
        catch
        {
            return UnavailableDescriptionLabelContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the description-label dropdown cannot be
    /// resolved: not available, carrying the previously published selection and
    /// options so a transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DescriptionLabelContext UnavailableDescriptionLabelContext(RibbonState previous) =>
        new(false, previous.SelectedDescriptionLabelPosition, previous.DescriptionLabelOptions);

    /// <summary>
    /// Captures the custom-style dropdown's context from the selection
    /// identifier, the table reader, and the catalogue reader (R5.6a).
    /// Never throws: an indeterminable selection keeps the previous context.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private static CustomStyleContext CaptureCustomStyleContext(
        SnapshotSources sources,
        RibbonState previous)
    {
        ISelectionIdentifier? identifier = sources.SelectionIdentifier;
        IGanttTableReader? reader = sources.TableReader;
        if (identifier is null || reader is null || sources.CatalogueReader is null)
        {
            return UnavailableCustomStyleContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            // Always capture the current catalogue of user-created styles,
            // independent of what row is selected. The dropdown's item list
            // is global to the workbook; only the enabled flag and selected
            // item depend on the selection. A failed catalogue read keeps
            // the previous options rather than presenting a stale-empty list.
            ConfigReadOutcome catalogues = sources.CatalogueReader.Read();
            IReadOnlyList<string> options = previous.CustomStyleOptions ?? [];
            if (catalogues.Succeeded && catalogues.Styles is not null)
            {
                HashSet<string> builtInKeys = new(StringComparer.Ordinal);
                foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
                {
                    _ = builtInKeys.Add(preset.StyleKey);
                }

                List<string> captured = [];
                foreach (string key in catalogues.Styles.StyleKeys)
                {
                    if (!builtInKeys.Contains(key))
                    {
                        captured.Add(key);
                    }
                }
                options = captured;
            }

            // Now resolve selection-dependent state: is the selection a single
            // style-key-type body row, and if so, which custom style does it
            // name?
            SelectionFacts facts = identifier.Identify();
            bool available = false;
            int bodyIndex = 0;
            if (facts.IsSingleBodyRow && facts.BodyIndex is int bi)
            {
                bodyIndex = bi;
                available = true;
            }
            string? selectedKey = null;

            if (available)
            {
                GanttTableReadOutcome read = reader.Read();
                if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
                {
                    available = false;
                }
                else
                {
                    GanttRowDto row = read.Rows[bodyIndex - 1];
                    if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type)
                        || EntityTypeCatalog.GetDefinition(type) is not { RequiresStyleKey: true })
                    {
                        available = false;
                    }
                    else
                    {
                        selectedKey = options.Contains(row.StyleKey ?? string.Empty, StringComparer.Ordinal)
                            ? row.StyleKey
                            : null;
                    }
                }
            }

            return new CustomStyleContext(available, selectedKey, options);
        }
        catch
        {
            return UnavailableCustomStyleContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the custom-style dropdown cannot be resolved:
    /// not available, carrying the previously published selection and options so
    /// a transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static CustomStyleContext UnavailableCustomStyleContext(RibbonState previous) =>
        new(false, previous.SelectedCustomStyleKey, previous.CustomStyleOptions);

    /// <summary>
    /// Captures the date-label dropdown's context from the selection identifier
    /// and the table reader (R5.14). The capability is type-only, so no
    /// catalogue or style registry is consulted. Never throws.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private static DateLabelContext CaptureDateLabelContext(SnapshotSources sources, RibbonState previous)
    {
        ISelectionIdentifier? identifier = sources.SelectionIdentifier;
        IGanttTableReader? reader = sources.TableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDateLabelContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDateLabelContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDateLabelContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type))
            {
                return UnavailableDateLabelContext(previous);
            }

            IReadOnlyList<GanttDateLabelPosition> options =
                EntityTypeCatalog.GetAllowedDateLabelPositions(type);
            if (options.Count == 0)
            {
                return UnavailableDateLabelContext(previous);
            }

            // The cell stores the choice by name; blank resolves to Auto
            // (R5.14 D3), so it reads as null.
            GanttDateLabelPosition? selected = EntityTypeCatalog.TryParseDateLabelSelection(
                row.DateLabelPositionText,
                out GanttDateLabelPosition? parsed)
                ? parsed
                : null;

            return new DateLabelContext(true, selected, options);
        }
        catch
        {
            return UnavailableDateLabelContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the date-label dropdown cannot be resolved:
    /// not available, carrying the previously published choice and options so a
    /// transient failure cannot blank a resolved dropdown.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DateLabelContext UnavailableDateLabelContext(RibbonState previous) =>
        new(false, previous.SelectedDateLabelPosition, previous.DateLabelOptions);

    /// <summary>
    /// Captures the date-label dropdown's context in the positionable vocabulary
    /// (owner ruling 2026-10-14): available only for the delineator, whose date
    /// label is positionable. Never throws.
    /// </summary>
    /// <param name="sources">The live inputs to read.</param>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The captured context, or the previous one when unavailable.</returns>
    private static DateLabelPlacementContext CaptureDateLabelPlacementContext(
        SnapshotSources sources,
        RibbonState previous)
    {
        ISelectionIdentifier? identifier = sources.SelectionIdentifier;
        IGanttTableReader? reader = sources.TableReader;
        if (identifier is null || reader is null)
        {
            return UnavailableDateLabelPlacementContext(previous);
        }

        // CA1031: the capture runs inside Ribbon callbacks and Excel events;
        // a failure keeps the last known context instead of propagating.
#pragma warning disable CA1031
        try
        {
            SelectionFacts facts = identifier.Identify();
            if (!facts.IsSingleBodyRow || facts.BodyIndex is not int bodyIndex)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            GanttTableReadOutcome read = reader.Read();
            if (!read.Succeeded || bodyIndex < 1 || bodyIndex > read.Rows.Count)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            GanttRowDto row = read.Rows[bodyIndex - 1];
            if (!EntityTypeCatalog.TryParse(row.TypeText, out GanttEntityType type)
                || type != GanttEntityType.Delineator)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            IReadOnlyList<GanttLabelPosition> options =
                EntityTypeCatalog.GetAllowedDateLabelPlacements(type);
            if (options.Count == 0)
            {
                return UnavailableDateLabelPlacementContext(previous);
            }

            // The cell stores the corner by name; blank resolves to Auto
            // (the which-date precedent), so it reads as null.
            GanttLabelPosition? selected = EntityTypeCatalog.TryParseDateLabelPlacementSelection(
                row.DateLabelPositionText,
                out GanttLabelPosition? parsed)
                ? parsed
                : null;

            return new DateLabelPlacementContext(true, selected, options);
        }
        catch
        {
            return UnavailableDateLabelPlacementContext(previous);
        }
#pragma warning restore CA1031
    }

    /// <summary>
    /// The context published when the positionable vocabulary cannot be resolved:
    /// not available, carrying the previously published corner and options.
    /// </summary>
    /// <param name="previous">The snapshot currently published.</param>
    /// <returns>The unavailable context.</returns>
    private static DateLabelPlacementContext UnavailableDateLabelPlacementContext(RibbonState previous) =>
        new(false, previous.SelectedDateLabelPlacement, previous.DateLabelPlacementOptions);

    /// <summary>
    /// The description-label dropdown's per-selection context.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored position, or null when blank/unknown.</param>
    /// <param name="Options">The row's allowed positions in display order.</param>
    private readonly record struct DescriptionLabelContext(
        bool Available,
        GanttLabelPosition? Selected,
        IReadOnlyList<GanttLabelPosition>? Options);

    /// <summary>
    /// The date-label dropdown's per-selection context.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored choice, or null when blank/unknown.</param>
    /// <param name="Options">The row's allowed choices in display order.</param>
    private readonly record struct DateLabelContext(
        bool Available,
        GanttDateLabelPosition? Selected,
        IReadOnlyList<GanttDateLabelPosition>? Options);

    /// <summary>
    /// The date-label dropdown's per-selection context in the positionable vocabulary.
    /// </summary>
    /// <param name="Available">True when exactly one body row is selected.</param>
    /// <param name="Selected">The row's stored corner, or null when blank (Auto).</param>
    /// <param name="Options">The row's allowed corners in display order.</param>
    private readonly record struct DateLabelPlacementContext(
        bool Available,
        GanttLabelPosition? Selected,
        IReadOnlyList<GanttLabelPosition>? Options);

    /// <summary>
    /// The custom-style dropdown's per-selection context (R5.6a).
    /// </summary>
    /// <param name="Available">True when exactly one style-key-type body row is selected.</param>
    /// <param name="SelectedKey">The row's stored custom style key, or null when blank/built-in.</param>
    /// <param name="Options">The user-created style keys in catalogue order.</param>
    private readonly record struct CustomStyleContext(
        bool Available,
        string? SelectedKey,
        IReadOnlyList<string>? Options);

    /// <summary>
    /// The structural row-action context published on the snapshot (R5.9b).
    /// </summary>
    /// <param name="Actions">The four availability flags for the active row.</param>
    /// <param name="ActiveCellInBodyRow">Whether the active cell resolves to a table body row.</param>
    private readonly record struct RowActionContext(
        RowActionAvailability Actions,
        bool ActiveCellInBodyRow);
}
