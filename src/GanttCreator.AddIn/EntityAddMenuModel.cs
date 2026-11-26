using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// The add-entity split buttons' model: the type-to-family mapping, each family's
/// default type, the ribbon id of its primary button, the glyph a type draws, and
/// the session's last menu pick per family.
/// </summary>
/// <remarks>
/// The remembered pick is session-scoped on purpose: persisting it is out of scope.
/// The permission question is asked through the same <c>getEnabled</c> truth table
/// the menu item uses, so the primary and the menu cannot disagree.
/// </remarks>
internal sealed class EntityAddMenuModel
{
    private readonly Func<string?, bool> _isMenuActionEnabled;

    /// <summary>
    /// The last entity added from each split button's menu, defaulting to the
    /// family's own default.
    /// </summary>
    private readonly Dictionary<RibbonStateService.AddEntityFamily, GanttEntityType> _lastAdded = new()
    {
        [RibbonStateService.AddEntityFamily.Activity] = GanttEntityType.AsPlannedActivity,
        [RibbonStateService.AddEntityFamily.Milestone] = GanttEntityType.AsPlannedMilestone,
        [RibbonStateService.AddEntityFamily.Element] = GanttEntityType.Delineator,
    };

    /// <summary>Creates the model.</summary>
    /// <param name="isMenuActionEnabled">
    /// The ribbon's own enablement truth table, asked per menu action id.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="isMenuActionEnabled"/> is null.</exception>
    internal EntityAddMenuModel(Func<string?, bool> isMenuActionEnabled)
    {
        _isMenuActionEnabled =
            isMenuActionEnabled ?? throw new ArgumentNullException(nameof(isMenuActionEnabled));
    }

    /// <summary>
    /// The type the family's primary button will add: the last menu pick, or the
    /// family default when that pick is not currently permitted.
    /// </summary>
    /// <param name="family">The split-button family.</param>
    /// <returns>The entity type the primary action would add right now.</returns>
    internal GanttEntityType PrimaryTypeFor(RibbonStateService.AddEntityFamily family)
    {
        GanttEntityType remembered = _lastAdded[family];
        return _isMenuActionEnabled(RibbonControlIds.EntityMenuItemPrefix + remembered.ToString())
            ? remembered
            : DefaultTypeFor(family);
    }

    /// <summary>Records a menu pick so that family's primary follows it.</summary>
    /// <param name="type">The type the menu action just added.</param>
    /// <returns>The family that recorded the pick, or null when no menu offers the type.</returns>
    internal RibbonStateService.AddEntityFamily? Remember(GanttEntityType type)
    {
        if (FamilyOf(type) is not { } family)
        {
            return null;
        }

        _lastAdded[family] = type;
        return family;
    }

    /// <summary>The family a type's menu belongs to, or <see langword="null"/> when no menu offers it.</summary>
    /// <param name="type">The entity type.</param>
    /// <returns>The family, or null for a type the ribbon cannot add from a split button.</returns>
    internal static RibbonStateService.AddEntityFamily? FamilyOf(GanttEntityType type) => type switch
    {
        GanttEntityType.AsPlannedActivity
        or GanttEntityType.AsBuiltActivity
        or GanttEntityType.BaselineActivity
        or GanttEntityType.CustomActivity
        or GanttEntityType.DelayEvent
        or GanttEntityType.CriticalInterval
        or GanttEntityType.AsPlannedProcurement
        or GanttEntityType.AsBuiltProcurement
        or GanttEntityType.BaselineProcurement => RibbonStateService.AddEntityFamily.Activity,

        GanttEntityType.AsPlannedMilestone
        or GanttEntityType.AsBuiltMilestone
        or GanttEntityType.BaselineMilestone
        or GanttEntityType.CriticalMilestone
        or GanttEntityType.CustomMilestone => RibbonStateService.AddEntityFamily.Milestone,

        GanttEntityType.Delineator
        or GanttEntityType.Spacer
        or GanttEntityType.Splitter => RibbonStateService.AddEntityFamily.Element,

        _ => null,
    };

    /// <summary>The family's default type, which is what its primary adds before any pick.</summary>
    /// <param name="family">The split-button family.</param>
    /// <returns>The default entity type.</returns>
    internal static GanttEntityType DefaultTypeFor(RibbonStateService.AddEntityFamily family) => family switch
    {
        RibbonStateService.AddEntityFamily.Activity => GanttEntityType.AsPlannedActivity,
        RibbonStateService.AddEntityFamily.Milestone => GanttEntityType.AsPlannedMilestone,
        RibbonStateService.AddEntityFamily.Element => GanttEntityType.Delineator,
        _ => GanttEntityType.AsPlannedActivity,
    };

    /// <summary>The ribbon id of a family's primary button.</summary>
    /// <param name="family">The split-button family.</param>
    /// <returns>The control id.</returns>
    internal static string PrimaryControlIdFor(RibbonStateService.AddEntityFamily family) => family switch
    {
        RibbonStateService.AddEntityFamily.Activity => RibbonControlIds.AddActivity,
        RibbonStateService.AddEntityFamily.Milestone => RibbonControlIds.AddMilestone,
        RibbonStateService.AddEntityFamily.Element => RibbonControlIds.AddDelineator,
        _ => RibbonControlIds.AddActivity,
    };

    /// <summary>The family a primary button id belongs to, or <see langword="null"/>.</summary>
    /// <param name="controlId">The control id.</param>
    /// <returns>The family, or null when the id is not one of the three primaries.</returns>
    internal static RibbonStateService.AddEntityFamily? FamilyOfPrimary(string? controlId) => controlId switch
    {
        RibbonControlIds.AddActivity => RibbonStateService.AddEntityFamily.Activity,
        RibbonControlIds.AddMilestone => RibbonStateService.AddEntityFamily.Milestone,
        RibbonControlIds.AddDelineator => RibbonStateService.AddEntityFamily.Element,
        _ => null,
    };

    /// <summary>
    /// The glyph slug for an add-entity primary, chosen so the picture agrees with
    /// the label beside it.
    /// </summary>
    /// <param name="type">The type the primary currently names.</param>
    /// <returns>The glyph slug.</returns>
    internal static string AddEntityPrimaryGlyph(GanttEntityType type) => type switch
    {
        GanttEntityType.Splitter => "splitter",
        GanttEntityType.Spacer => "spacer",
        GanttEntityType.Delineator => "delineator",

        // R5.16a surface revision 2026-11-17: the three concept rows the menus draw
        // differently from their family. The style-defined pair uses the draft ring
        // rather than a pen: at 16 px the pen's nib read as a padlock on the owner's
        // live review, and a padlock on this tab means the cells or the chart are
        // locked, so the wrong reading is a state contradiction rather than a
        // near-miss.
        GanttEntityType.AsPlannedProcurement
            or GanttEntityType.AsBuiltProcurement
            or GanttEntityType.BaselineProcurement => "procurement",
        GanttEntityType.DelayEvent => "span-delay",
        GanttEntityType.CustomActivity
            or GanttEntityType.CustomMilestone => "circle-dashed",

        // The critical pair keeps its family shape and gains the shared slash, so the
        // primary names "critical" without lying about whether it is a bar or a flag.
        GanttEntityType.CriticalInterval => "critical-span",
        GanttEntityType.CriticalMilestone => "critical-milestone",

        GanttEntityType.AsBuiltMilestone
            or GanttEntityType.AsPlannedMilestone
            or GanttEntityType.BaselineMilestone => "add-milestone",
        GanttEntityType.AsBuiltActivity
            or GanttEntityType.AsPlannedActivity
            or GanttEntityType.BaselineActivity => "add-activity",

        // An unnamed value cannot reach here — the remembered type always comes from the
        // catalogue — but the enum is not a closed set to the compiler, so it needs an arm.
        _ => "add-activity",
    };
}
