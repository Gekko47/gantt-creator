namespace GanttCreator.Core;

/// <summary>
/// The identity state of the add-in configuration worksheet. The state decides
/// whether the recovery path may create a sheet (never beside a damaged
/// original), may rename one back, or must refuse.
/// </summary>
public enum ConfigSheetState
{
    /// <summary>Exactly one worksheet carries the contract name, or is a stripped VeryHidden one.</summary>
    Present = 0,

    /// <summary>No worksheet can plausibly be the configuration; creation is allowed.</summary>
    Missing = 1,

    /// <summary>Exactly one non-canonical VeryHidden worksheet carries contract tables; rename it back.</summary>
    Renamed = 2,

    /// <summary>A visible worksheet carries the contract name but no contract table; never touch it.</summary>
    NameCollision = 3,

    /// <summary>More than one worksheet carries the contract name, or more than one renamed candidate.</summary>
    Ambiguous = 4,

    /// <summary>A visible worksheet carries a contract table; refuse rather than create a second configuration.</summary>
    AmbiguousCandidate = 5,

    /// <summary>A non-canonical VeryHidden worksheet carries no contract table; it may be a stripped configuration.</summary>
    Unrecognised = 6,
}

/// <summary>
/// One worksheet as the configuration locator sees it. Office-free so the
/// identity rules are testable without a host.
/// </summary>
/// <param name="Name">The live worksheet name.</param>
/// <param name="IsVeryHidden">Whether the worksheet visibility is <c>xlSheetVeryHidden</c>.</param>
/// <param name="ContractTableCount">How many contract tables the worksheet carries.</param>
public sealed record ConfigSheetObservation(string Name, bool IsVeryHidden, int ContractTableCount);

/// <summary>The result of locating the configuration worksheet.</summary>
/// <param name="State">The identity state.</param>
/// <param name="SheetName">The resolved worksheet name, when the state identifies one.</param>
public sealed record ConfigSheetLocation(ConfigSheetState State, string? SheetName)
{
    /// <summary>Whether the recovery renames a misnamed configuration sheet back to the contract name.</summary>
    public bool NeedsRename => State == ConfigSheetState.Renamed;
}

/// <summary>
/// The single authority for which worksheet is the configuration, and whether
/// it may be created or renamed.
/// </summary>
/// <remarks>
/// <para>
/// Creating a configuration worksheet is only allowed in
/// <see cref="ConfigSheetState.Missing"/>. Every other non-present state either
/// renames the one candidate back or refuses, because a second helper sheet
/// breaks the workbook invariant.
/// </para>
/// <para>
/// A canonical name plus <c>xlSheetVeryHidden</c> is our sheet even when every
/// catalogue table has been deleted: a user cannot create that visibility from
/// the ordinary Excel UI. A canonical name that is not VeryHidden and carries no
/// contract table is a user sheet that merely took the name, and is never
/// hidden, written, or deleted.
/// </para>
/// </remarks>
public static class ConfigSheetLocator
{
    /// <summary>
    /// Locates the configuration worksheet from the observed worksheets.
    /// </summary>
    /// <param name="sheets">Every worksheet in the workbook, in sheet order.</param>
    /// <returns>The location and its identity state.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="sheets"/> is null.</exception>
    public static ConfigSheetLocation Locate(IReadOnlyList<ConfigSheetObservation> sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        ConfigSheetObservation? canonical = null;
        int canonicalCount = 0;
        foreach (ConfigSheetObservation sheet in sheets)
        {
            if (!IsCanonicalName(sheet.Name))
            {
                continue;
            }

            canonical = sheet;
            canonicalCount++;
        }

        if (canonicalCount > 1)
        {
            return new ConfigSheetLocation(ConfigSheetState.Ambiguous, null);
        }

        if (canonical is not null)
        {
            return canonical.ContractTableCount > 0 || canonical.IsVeryHidden
                ? new ConfigSheetLocation(ConfigSheetState.Present, canonical.Name)
                : new ConfigSheetLocation(ConfigSheetState.NameCollision, canonical.Name);
        }

        ConfigSheetObservation? renamed = null;
        int renamedCount = 0;
        foreach (ConfigSheetObservation sheet in sheets)
        {
            if (sheet.IsVeryHidden && sheet.ContractTableCount > 0)
            {
                renamed = sheet;
                renamedCount++;
            }
        }

        if (renamedCount > 1)
        {
            return new ConfigSheetLocation(ConfigSheetState.Ambiguous, null);
        }

        if (renamed is not null)
        {
            return new ConfigSheetLocation(ConfigSheetState.Renamed, renamed.Name);
        }

        foreach (ConfigSheetObservation sheet in sheets)
        {
            if (!sheet.IsVeryHidden && sheet.ContractTableCount > 0)
            {
                return new ConfigSheetLocation(ConfigSheetState.AmbiguousCandidate, null);
            }
        }

        foreach (ConfigSheetObservation sheet in sheets)
        {
            if (sheet.IsVeryHidden)
            {
                return new ConfigSheetLocation(ConfigSheetState.Unrecognised, null);
            }
        }

        return new ConfigSheetLocation(ConfigSheetState.Missing, null);
    }

    private static bool IsCanonicalName(string? name) =>
        string.Equals(name, GanttWorkbookContract.ConfigSheetName, StringComparison.OrdinalIgnoreCase);
}
