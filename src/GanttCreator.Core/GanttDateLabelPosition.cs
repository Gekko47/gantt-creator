namespace GanttCreator.Core;

/// <summary>
/// Which of a row's section-23 date labels render. Distinct from
/// <see cref="GanttLabelPosition"/>, which describes <em>where</em> a
/// label sits: this selects <em>which</em> date label (start, finish,
/// or both) the row shows. Blank <c>DateLabelPosition</c> resolves to
/// <see cref="Auto"/>; the permitted set per type is derived by
/// <see cref="EntityTypeCatalog.GetAllowedDateLabelPositions"/>.
/// </summary>
/// <remarks>
/// <para>
/// The set a type may name is derived from
/// <see cref="EntityTypeDefinition.DateMode"/>, never authored as a
/// second list: a dateless type names none (the control is disabled),
/// a <c>StartOnly</c> type names <c>Auto</c> and <c>Start</c>, and a
/// span names all three. <c>Finish</c> is therefore never offered for a
/// <c>StartOnly</c> type, which reads no finish date.
/// </para>
/// </remarks>
public enum GanttDateLabelPosition
{
    /// <summary>Both date labels the row's type uses (the default).</summary>
    Auto = 0,

    /// <summary>The start-date label only.</summary>
    Start = 1,

    /// <summary>The finish-date label only (span types only).</summary>
    Finish = 2,
}
