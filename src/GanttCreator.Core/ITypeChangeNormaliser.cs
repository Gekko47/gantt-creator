namespace GanttCreator.Core;

/// <summary>
/// Outcome of <see cref="ITypeChangeNormaliser.Normalise"/>: how many engine-hidden
/// cells were corrected to match each row's current <c>Type</c>.
/// </summary>
/// <param name="CellsWritten">Number of engine-hidden cells rewritten.</param>
public sealed record GanttTypeChangeNormaliseOutcome(int CellsWritten)
{
    /// <summary>Successful outcome with no cells written.</summary>
    public static GanttTypeChangeNormaliseOutcome Ok(int cellsWritten) => new(cellsWritten);
}

/// <summary>
/// Normalises type-dependent engine-hidden columns so a row's <c>Type</c> change
/// does not leave stale values that the validator would reject. Pure and Office-free:
/// computes corrections from the DTO values; the Office adapter writes them.
/// </summary>
public interface ITypeChangeNormaliser
{
    /// <summary>Normalises engine-hidden columns for every row and returns how many cells were written.</summary>
    GanttTypeChangeNormaliseOutcome Normalise();
}