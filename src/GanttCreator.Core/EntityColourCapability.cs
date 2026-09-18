namespace GanttCreator.Core;

/// <summary>
/// The colour-override capabilities of a row type, defining which per-row
/// colour columns are valid for it. The entity guide's wording maps as:
/// "Fill" → <see cref="Fill"/>; "Fill + outline" → <see cref="Fill"/> |
/// <see cref="Stroke"/>; "Hatch + outline" → <see cref="Hatch"/> |
/// <see cref="Stroke"/>; "Stroke" → <see cref="Stroke"/>; "none" →
/// <see cref="None"/>; "style-defined" → <see cref="StyleDefined"/>.
/// </summary>
[Flags]
public enum EntityColourCapability
{
    /// <summary>No colour overrides are valid for the type.</summary>
    None = 0,

    /// <summary>The row may override the rectangle fill colour (<c>FillColour</c>).</summary>
    Fill = 1,

    /// <summary>The row may override the outline or line colour (<c>StrokeColour</c>).</summary>
    Stroke = 2,

    /// <summary>The row renders a hatched fill that may be overridden.</summary>
    Hatch = 4,

    /// <summary>The colour scheme is defined entirely by the row's named style.</summary>
    StyleDefined = 8,
}
