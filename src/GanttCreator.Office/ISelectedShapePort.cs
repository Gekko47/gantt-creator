namespace GanttCreator.Office;

/// <summary>Why the current selection is not one shape this add-in owns.</summary>
/// <remarks>
/// Typed and expected: a user who has a cell selected, or a shape the add-in did not draw, is not an
/// error. The Ribbon reads these to decide whether the Unlock button is enabled, so none of them may
/// throw.
/// </remarks>
public enum SelectedShapeRefusalReason
{
    /// <summary>The Excel application object was absent.</summary>
    NoApplication = 0,

    /// <summary>Nothing is selected, or the selection is not a shape (a cell range, a chart).</summary>
    SelectionIsNotASingleShape = 1,

    /// <summary>The selected shape is not one this add-in drew, so it is never unlocked.</summary>
    ShapeNotOwned = 2,

    /// <summary>The host refused a property read while resolving the selection.</summary>
    HostRejected = 3,
}

/// <summary>The selection resolved to one owned shape, or a typed refusal.</summary>
/// <param name="PrimitiveId">The scene primitive id of the selected owned shape, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/> on success.</param>
public sealed record SelectedShapeOutcome(string? PrimitiveId, SelectedShapeRefusalReason? Refusal)
{
    /// <summary>Gets whether one owned shape is selected.</summary>
    public bool Succeeded => PrimitiveId is not null && Refusal is null;

    /// <summary>Creates a success outcome.</summary>
    /// <param name="primitiveId">The selected owned shape's identifier.</param>
    /// <returns>The success outcome.</returns>
    public static SelectedShapeOutcome Ok(string primitiveId) => new(primitiveId, null);

    /// <summary>Creates a refusal.</summary>
    /// <param name="refusal">The refusal reason.</param>
    /// <returns>The refusal outcome.</returns>
    public static SelectedShapeOutcome Refused(SelectedShapeRefusalReason refusal) => new(null, refusal);
}

/// <summary>
/// Resolves the user's current selection to one shape this add-in owns (ADR-0055 clause 6).
/// </summary>
/// <remarks>
/// <para>
/// The Unlock command acts only on a selection that is provably ours, because unlocking is the one
/// operation that hands the user a writable chart element. Ownership is proved by the ownership tag
/// recomputed from the shape's name, the same rule every write operation follows, so a user's shape
/// that happens to share a name is refused.
/// </para>
/// <para>
/// No interop type crosses this port, so the AddIn compilation never names
/// <c>Microsoft.Office.Interop.Excel</c>.
/// </para>
/// </remarks>
public interface ISelectedShapePort
{
    /// <summary>Resolves the current selection.</summary>
    /// <returns>The selected owned shape, or a typed refusal. Never throws.</returns>
    SelectedShapeOutcome Read();
}
