namespace GanttCreator.Core.Scene;

/// <summary>
/// The composition profile a scene is being built for, which decides whether a
/// <em>data panel</em> — a drawn replica of the visible worksheet table — belongs
/// in the output at all (R4.8A D4).
/// </summary>
/// <remarks>
/// <para>
/// <b>Why this is a type and not a convention.</b> <see cref="SceneBuildRequest.Panel"/>
/// is a nullable <see cref="PanelTheme"/>, so "Live must not draw a panel" was
/// already <em>expressible</em> before this type existed — a caller simply left it
/// null. But expressible is not enforced: a factory that set <c>Panel</c> for a live
/// refresh would compile, run, and cover the user's own worksheet cells with a drawn
/// replica of them. The failure is silent and looks like a rendering bug, and on the
/// live sheet it destroys the thing the user is editing.
/// </para>
/// <para>
/// The profile makes the decision <b>structural</b>. A profile is chosen in code
/// rather than by a caller remembering which fields to leave null, and the
/// expectation is asserted rather than documented. That is the whole point of D4's
/// "chosen in code, not by a caller remembering to pass the right theme".
/// </para>
/// <para>
/// <b>Why Live is the odd one out.</b> Every other profile renders into a destination
/// that has no table behind it — a PNG, a PowerPoint slide, a copy the user is about
/// to edit somewhere else — so a panel is the only representation of the row labels
/// and they would otherwise be missing entirely. The live sheet is the opposite: the
/// real cells <em>are</em> the panel, sitting directly beside the plot, so drawing a
/// second one is not a missing feature but a duplicated, overlapping one.
/// </para>
/// </remarks>
public enum SceneCompositionProfile
{
    /// <summary>
    /// The live Excel worksheet. The real cells are the data panel, so none is drawn.
    /// </summary>
    LiveExcel = 0,

    /// <summary>An editable PowerPoint or Excel copy, which has no cells of its own.</summary>
    EditableExport = 1,

    /// <summary>A PowerPoint slide, which has no cells of its own.</summary>
    PowerPoint = 2,

    /// <summary>A flat raster image, which has no cells of its own.</summary>
    Raster = 3,
}

/// <summary>
/// The composition rules a <see cref="SceneCompositionProfile"/> implies, so the
/// "does this profile draw a data panel" decision is stated once.
/// </summary>
/// <remarks>
/// A pure function of the profile, with no dependency on a request, a workbook, or
/// Excel. That is deliberate: the rule must be testable on its own, because a
/// violation of it is silent in production and only visible as a wrong picture.
/// </remarks>
public static class SceneCompositionProfiles
{
    /// <summary>
    /// Determines whether this profile draws a data panel into its output.
    /// </summary>
    /// <param name="profile">The composition profile.</param>
    /// <returns>
    /// <see langword="false"/> for <see cref="SceneCompositionProfile.LiveExcel"/>,
    /// because the worksheet's own cells are the panel; <see langword="true"/> for
    /// every other profile, whose destination carries no cells.
    /// </returns>
    public static bool DrawsDataPanel(SceneCompositionProfile profile) => profile switch
    {
        // The only profile whose destination already contains the table.
        SceneCompositionProfile.LiveExcel => false,
        SceneCompositionProfile.EditableExport => true,
        SceneCompositionProfile.PowerPoint => true,
        SceneCompositionProfile.Raster => true,

        // An unrecognised value must NOT default to "draws a panel". A new profile
        // added to the enum without a rule here would otherwise silently receive a
        // drawn panel, and on the live sheet that is the destructive case. Failing
        // closed keeps the safe behaviour the default until someone decides.
        _ => false,
    };
}
