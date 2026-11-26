using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// Pure projections of <see cref="RibbonState"/> into the three label dropdowns'
/// item counts, item IDs, user-facing labels and selected indexes.
/// </summary>
/// <remarks>
/// Every member is a pure read of the cached snapshot: no Excel probe, no ribbon
/// call, no log write, and no exception for any input. The wording tables live
/// here so each display string has exactly one authority, and the Ribbon getters
/// in <see cref="RibbonStateService"/> stay one-line delegations.
/// </remarks>
internal static class RibbonDropdownProjection
{
    /// <summary>The leading clear item's ID for the custom-style dropdown.</summary>
    internal const string NoneItem = "None";

    /// <summary>The leading clear item's user-facing label.</summary>
    internal const string NoneLabel = "None";

    /// <summary>
    /// Returns how many items the description-label dropdown offers: the row's
    /// projected positions plus the leading <c>Default</c> item that clears the
    /// cell (R5.5 ruling 5: blank resolves to the style default / Auto cascade).
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>The item count, or 0 when no single row is selected.</returns>
    internal static int GetDescriptionLabelCount(RibbonState state) =>
        state.DescriptionLabelOptions is { Count: > 0 } options ? options.Count + 1 : 0;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: <c>Default</c> for the
    /// leading clear item, otherwise the position's enum name.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The item ID, or an empty string for an out-of-range index.</returns>
    internal static string GetDescriptionLabelItemID(RibbonState state, int index) =>
        index == 0
            ? EntityTypeCatalog.DefaultLabelSelectionId
            : state.DescriptionLabelOptions is { Count: > 0 } options
            && index >= 1
            && index <= options.Count
                ? options[index - 1].ToString()
                : string.Empty;

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/> (the enum name is
    /// never shown verbatim; <see cref="LabelPositionDisplayText"/> is the single
    /// wording authority).
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The label, or an empty string for an out-of-range index.</returns>
    internal static string GetDescriptionLabelItemLabel(RibbonState state, int index) =>
        index == 0
            ? "Default"
            : state.DescriptionLabelOptions is { Count: > 0 } options
            && index >= 1
            && index <= options.Count
                ? LabelPositionDisplayText(options[index - 1])
                : string.Empty;

    /// <summary>
    /// Returns the selected item index: 0 (<c>Default</c>) when the row's cell is
    /// blank, otherwise one past the position's place in the projected options.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>The selected index, never negative.</returns>
    internal static int GetDescriptionLabelSelectedIndex(RibbonState state)
    {
        if (state.DescriptionLabelOptions is not { Count: > 0 } options)
        {
            return 0;
        }

        if (state.SelectedDescriptionLabelPosition is not GanttLabelPosition selected)
        {
            return 0;
        }

        for (int index = 0; index < options.Count; index++)
        {
            if (options[index] == selected)
            {
                return index + 1;
            }
        }

        return 0;
    }

    /// <summary>
    /// Returns the display text for a label position. Isolated so the wording
    /// table has exactly one authority.
    /// </summary>
    /// <param name="position">The position being displayed.</param>
    /// <returns>The user-facing wording.</returns>
    internal static string LabelPositionDisplayText(GanttLabelPosition position) =>
        position switch
        {
            GanttLabelPosition.Auto => "Auto",
            GanttLabelPosition.Left => "Left",
            GanttLabelPosition.Right => "Right",
            GanttLabelPosition.Inside => "Centre",
            GanttLabelPosition.TopLeft => "Top left",
            GanttLabelPosition.TopRight => "Top right",
            GanttLabelPosition.BottomLeft => "Bottom left",
            GanttLabelPosition.BottomRight => "Bottom right",
            GanttLabelPosition.DataPanelLeft => "Data panel",
            GanttLabelPosition.PlotCentre => "Plot centre",
            GanttLabelPosition.Both => "Both",
            GanttLabelPosition.None => "None (hide)",
            _ => position.ToString(),
        };

    /// <summary>
    /// Returns how many items the custom-style dropdown offers: a leading
    /// <c>None</c> clear item plus the user-created style keys.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>Always at least 1, so Excel never queries an out-of-range index.</returns>
    internal static int GetCustomStyleCount(RibbonState state) =>
        (state.CustomStyleOptions is { Count: > 0 } options ? options.Count : 0) + 1;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: <c>None</c> for the
    /// leading clear item, otherwise the style key itself.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The item ID, or an empty string for an out-of-range index.</returns>
    internal static string GetCustomStyleItemID(RibbonState state, int index) =>
        index == 0 ? NoneItem :
        state.CustomStyleOptions is { Count: > 0 } options
        && index >= 1
        && index < options.Count + 1
            ? options[index - 1]
            : string.Empty;

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/>: <c>None</c> for
    /// the clear item, otherwise the style key shown verbatim.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The label, or an empty string for an out-of-range index.</returns>
    internal static string GetCustomStyleItemLabel(RibbonState state, int index) =>
        index == 0 ? NoneLabel : GetCustomStyleItemID(state, index);

    /// <summary>
    /// Returns the selected item index: 0 for <c>None</c> when the cell is blank,
    /// names a built-in preset, or no single style-key row is selected.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>The selected index, never negative.</returns>
    internal static int GetCustomStyleSelectedIndex(RibbonState state)
    {
        if (state.CustomStyleOptions is not { Count: > 0 } options
            || state.SelectedCustomStyleKey is not string selected)
        {
            return 0;
        }

        for (int index = 0; index < options.Count; index++)
        {
            if (string.Equals(options[index], selected, StringComparison.Ordinal))
            {
                return index + 1;
            }
        }

        return 0;
    }

    /// <summary>
    /// Returns how many items the date-label dropdown offers: the selected row's
    /// allowed date-label choices, projected from its <c>Type</c> (R5.14 D1/D4).
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>The item count, or 0 when no single row is selected.</returns>
    internal static int GetDateLabelCount(RibbonState state) =>
        state.DateLabelPlacementSelectionAvailable
        && state.DateLabelPlacementOptions is { Count: > 0 } placementOptions
            ? placementOptions.Count
            : state.DateLabelOptions is { Count: > 0 } options ? options.Count : 0;

    /// <summary>
    /// Returns the item ID at <paramref name="index"/>: the leading <c>Auto</c>
    /// choice carries the <c>Default</c> clear-ID (selecting it clears the cell to
    /// blank = Auto per R5.14 D3), and every other choice carries its enum name.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The item ID, or an empty string for an out-of-range index.</returns>
    internal static string GetDateLabelItemID(RibbonState state, int index)
    {
        // Owner ruling 2026-10-14: a delineator row carries the
        // positionable vocabulary, whose items are the description
        // label's corners. Auto clears the cell, so its item id is
        // the clear marker, not the enum name
        // (TryParseDateLabelPlacementSelection maps Default to null).
        if (state.DateLabelPlacementSelectionAvailable
            && state.DateLabelPlacementOptions is { Count: > 0 } placementOptions)
        {
            return index < 0 || index >= placementOptions.Count
                ? string.Empty
                : placementOptions[index] == GanttLabelPosition.Auto
                    ? EntityTypeCatalog.DefaultDateLabelSelectionId
                    : placementOptions[index].ToString();
        }

        if (state.DateLabelOptions is not { Count: > 0 } options
            || index < 0 || index >= options.Count)
        {
            return string.Empty;
        }

        // Auto clears the cell, so its item id is the clear marker, not
        // the enum name (TryParseDateLabelSelection maps Default to null).
        return options[index] == GanttDateLabelPosition.Auto
            ? EntityTypeCatalog.DefaultDateLabelSelectionId
            : options[index].ToString();
    }

    /// <summary>
    /// Returns the user-facing label at <paramref name="index"/> (the enum name is
    /// never shown verbatim; <see cref="DateLabelDisplayText"/> and
    /// <see cref="LabelPositionDisplayText"/> are the wording authorities).
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <param name="index">The zero-based item index.</param>
    /// <returns>The label, or an empty string for an out-of-range index.</returns>
    internal static string GetDateLabelItemLabel(RibbonState state, int index) =>
        state.DateLabelPlacementSelectionAvailable
        && state.DateLabelPlacementOptions is { Count: > 0 } placementOptions
        && index >= 0
        && index < placementOptions.Count
            ? LabelPositionDisplayText(placementOptions[index])
            : state.DateLabelOptions is { Count: > 0 } options
            && index >= 0
            && index < options.Count
                ? DateLabelDisplayText(options[index])
                : string.Empty;

    /// <summary>
    /// Returns the selected item index: <c>Auto</c> (index 0) when the row's cell
    /// is blank, otherwise the choice's place in the projected options.
    /// </summary>
    /// <param name="state">The snapshot being projected.</param>
    /// <returns>The selected index, never negative.</returns>
    internal static int GetDateLabelSelectedIndex(RibbonState state)
    {
        // Owner ruling 2026-10-14: the positionable vocabulary's
        // leading choice is Auto, exactly as the which-date vocabulary's
        // is, so a blank cell reads as index 0 in either.
        if (state.DateLabelPlacementSelectionAvailable
            && state.DateLabelPlacementOptions is { Count: > 0 } placementOptions)
        {
            // Blank (null) resolves to Auto, and Auto is the leading choice.
            GanttLabelPosition selectedPlacement = state.SelectedDateLabelPlacement ?? GanttLabelPosition.Auto;
            for (int index = 0; index < placementOptions.Count; index++)
            {
                if (placementOptions[index] == selectedPlacement)
                {
                    return index;
                }
            }

            return 0;
        }

        if (state.DateLabelOptions is not { Count: > 0 } options)
        {
            return 0;
        }

        // Blank (null) resolves to Auto, and Auto is the leading choice.
        GanttDateLabelPosition selected = state.SelectedDateLabelPosition ?? GanttDateLabelPosition.Auto;
        for (int index = 0; index < options.Count; index++)
        {
            if (options[index] == selected)
            {
                return index;
            }
        }

        return 0;
    }

    /// <summary>
    /// Returns the display text for a date-label choice. Isolated so the wording
    /// table has exactly one authority.
    /// </summary>
    /// <param name="position">The choice being displayed.</param>
    /// <returns>The user-facing wording.</returns>
    internal static string DateLabelDisplayText(GanttDateLabelPosition position) =>
        position switch
        {
            // R5.18m (owner ruling 2026-11-13): the leading item clears the cell,
            // and a blank cell resolves to the type default (None), so the item is
            // named for what it does rather than for the legacy stored value.
            GanttDateLabelPosition.Auto => "Default",
            GanttDateLabelPosition.Both => "Both",
            GanttDateLabelPosition.Start => "Start",
            GanttDateLabelPosition.Finish => "Finish",
            GanttDateLabelPosition.None => "None",
            _ => position.ToString(),
        };
}
