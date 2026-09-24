namespace GanttCreator.Core.Scene;

/// <summary>An immutable text primitive with already-resolved bounds and typography.</summary>
public sealed record SceneText : ScenePrimitive
{
    /// <summary>Initialises text and validates its content.</summary>
    /// <param name="primitiveId">The stable role-derived primitive identifier.</param>
    /// <param name="ownerId">The stable source row identifier.</param>
    /// <param name="zLayer">The scene layer.</param>
    /// <param name="text">The text content.</param>
    /// <param name="textBounds">The resolved text bounds in points.</param>
    /// <param name="style">The resolved style.</param>
    /// <param name="alignment">The resolved text alignment.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    public SceneText(
        string primitiveId,
        GanttRowId ownerId,
        ZLayer zLayer,
        string text,
        RectD textBounds,
        SceneStyle style,
        GanttLabelPosition alignment,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null)
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(style);
        if (!Enum.IsDefined(alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }

        Text = text;
        TextBounds = textBounds;
        Style = style;
        Alignment = alignment;
    }

    /// <summary>Gets the text content.</summary>
    public string Text { get; }

    /// <summary>Gets the resolved text bounds in points.</summary>
    public RectD TextBounds { get; }

    /// <summary>Gets the resolved style.</summary>
    public SceneStyle Style { get; }

    /// <summary>Gets the resolved text alignment.</summary>
    public GanttLabelPosition Alignment { get; }
}
