namespace GanttCreator.Core.Scene;

/// <summary>
/// Owns the role prefixes <see cref="PanelBuilder"/> stamps on the section 3/4
/// data-panel primitives, and the single predicate that identifies one.
/// </summary>
/// <remarks>
/// <para>
/// <b>Why a predicate and not a flag on the primitive.</b> R6.2's editable-export
/// composition stages the data panel out of a mixed export scene, so it must answer
/// "which primitives are the panel's?" without inventing a parallel model - the scene
/// is the single source and renderers consume it. The answer is a property of the role
/// the panel builder already stamps, so it is stated once here, beside the emitter,
/// rather than restated as string literals in the Office composition.
/// </para>
/// <para>
/// <b>The prefixes are the emitter's contract.</b> A role that <see cref="PanelBuilder"/>
/// adds or renames without a matching prefix here would silently drop that family from
/// an export. <c>PanelPrimitiveRolesTests</c> pins every primitive the builder emits
/// against this predicate - including a check that all six families occur - so that
/// drift fails in the gate rather than appearing as a missing cell in an export.
/// </para>
/// </remarks>
public static class PanelPrimitiveRoles
{
    private const string _headerCellPrefix = "header-cell:";
    private const string _headerTextPrefix = "header-text:";
    private const string _panelCellPrefix = "panel-cell:";
    private const string _panelTextPrefix = "panel-text:";
    private const string _panelBorderVerticalPrefix = "panel-border-v:";
    private const string _panelBorderHorizontalPrefix = "panel-border-h:";

    /// <summary>Determines whether a primitive belongs to the section 3/4 data panel.</summary>
    /// <param name="primitive">The primitive to classify, or <see langword="null"/>.</param>
    /// <returns>
    /// <see langword="true"/> when the primitive's role is one the panel builder stamps;
    /// otherwise <see langword="false"/>.
    /// </returns>
    /// <remarks>
    /// The role is the primitive identifier after the owner value and its colon,
    /// exactly as <see cref="ScenePrimitive.CreateId"/> composes it. An identifier that
    /// does not carry its own owner prefix is treated as foreign and answers
    /// <see langword="false"/>: classification fails closed, and a substring test would
    /// match a row identifier that merely contains a prefix somewhere inside it.
    /// </remarks>
    public static bool IsPanelPrimitive(ScenePrimitive? primitive)
    {
        if (primitive is null)
        {
            return false;
        }

        string ownerPrefix = primitive.OwnerId.Value + ":";
        if (!primitive.PrimitiveId.StartsWith(ownerPrefix, StringComparison.Ordinal))
        {
            return false;
        }

        ReadOnlySpan<char> role = primitive.PrimitiveId.AsSpan(ownerPrefix.Length);
        return role.StartsWith(_headerCellPrefix, StringComparison.Ordinal)
            || role.StartsWith(_headerTextPrefix, StringComparison.Ordinal)
            || role.StartsWith(_panelCellPrefix, StringComparison.Ordinal)
            || role.StartsWith(_panelTextPrefix, StringComparison.Ordinal)
            || role.StartsWith(_panelBorderVerticalPrefix, StringComparison.Ordinal)
            || role.StartsWith(_panelBorderHorizontalPrefix, StringComparison.Ordinal);
    }
}
