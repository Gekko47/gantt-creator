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
    /// <param name="style">
    /// The resolved style. A style that names no font family is stored through
    /// <see cref="SceneStyle.WithLabelFontFamilyIfMissing"/>: the scene, not the renderer,
    /// names the font, so a text primitive can never reach a renderer without one
    /// (R8.8 live defect, 2026-11-23).
    /// </param>
    /// <param name="alignment">The resolved text alignment.</param>
    /// <param name="entityType">The source entity type, when known.</param>
    /// <param name="laneOrder">The lane ordering value, when known.</param>
    /// <param name="stackIndex">The stack ordering value, when known.</param>
    /// <param name="sortOrder">The explicit user ordering value, when known.</param>
    /// <param name="leftInsetPt">
    /// The left inset inside <paramref name="textBounds"/> the renderer applies before the
    /// text starts, in points. Zero (the default) means the text begins at the bounds' left
    /// edge. A grouped child's exported <c>Description</c> cell sets this to reproduce the
    /// worksheet cell's <c>IndentLevel</c> as a text-frame margin.
    /// </param>
    public SceneText(
        string primitiveId,
        SceneOwnerId ownerId,
        ZLayer zLayer,
        string text,
        RectD textBounds,
        SceneStyle style,
        GanttTextAlignment alignment,
        GanttEntityType? entityType = null,
        int? laneOrder = null,
        int? stackIndex = null,
        int? sortOrder = null,
        double leftInsetPt = 0
    )
        : base(primitiveId, ownerId, zLayer, entityType, laneOrder, stackIndex, sortOrder)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(style);
        if (!Enum.IsDefined(alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }

        if (!double.IsFinite(leftInsetPt) || leftInsetPt < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(leftInsetPt));
        }

        Text = text;
        TextBounds = textBounds;
        Style = style.WithLabelFontFamilyIfMissing();
        Alignment = alignment;
        LeftInsetPt = leftInsetPt;
    }

    /// <summary>Gets the text content.</summary>
    public string Text { get; }

    /// <summary>Gets the resolved text bounds in points.</summary>
    public RectD TextBounds { get; }

    /// <summary>Gets the resolved style.</summary>
    public SceneStyle Style { get; }

    /// <summary>Gets the resolved text alignment.</summary>
    public GanttTextAlignment Alignment { get; }

    /// <summary>
    /// Gets the left inset the renderer applies inside <see cref="TextBounds"/> before the
    /// text starts, in points.
    /// </summary>
    /// <remarks>
    /// <see cref="SceneText"/> is the one primitive the data-panel export reuses for a
    /// worksheet cell's text. The live sheet expresses a grouped child's indent as a cell
    /// <c>IndentLevel</c>; the copy draws shapes, so the indent must travel on the primitive.
    /// Zero is the default and leaves every existing text primitive byte-identical.
    /// </remarks>
    public double LeftInsetPt { get; }
}
