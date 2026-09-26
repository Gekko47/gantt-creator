namespace GanttCreator.Core;

/// <summary>
/// The horizontal alignment of a text run inside its already-resolved text
/// bounds.
/// </summary>
/// <remarks>
/// <para>
/// This is deliberately a separate type from <see cref="GanttLabelPosition"/>.
/// The two answer different questions: a label position says <em>where the
/// label box sits</em> relative to its entity, while a text alignment says
/// <em>how the text sits inside the box the scene already resolved</em>. A
/// single enum covering both cannot express §2's centred chart title, because
/// <c>Centre</c> is not a label position.
/// </para>
/// <para>
/// The scene always resolves the box, so a renderer consumes this value and
/// never re-measures or re-chooses a side (ADR-0013, ADR-0015).
/// </para>
/// </remarks>
public enum GanttTextAlignment
{
    /// <summary>Text starts at the left edge of its resolved bounds.</summary>
    Left = 0,

    /// <summary>Text is centred within its resolved bounds.</summary>
    Centre = 1,

    /// <summary>Text ends at the right edge of its resolved bounds.</summary>
    Right = 2,
}
