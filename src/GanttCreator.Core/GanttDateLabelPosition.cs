namespace GanttCreator.Core;

/// <summary>
/// Which of a row's section-23 date labels render. Distinct from
/// <see cref="GanttLabelPosition"/>, which describes <em>where</em> a
/// label sits: this selects <em>which</em> date label (start, finish,
/// both, or none) the row shows. A blank <c>DateLabelPosition</c>
/// resolves to the <em>type default</em>, which is
/// <see cref="None"/> for every date-bearing type; the permitted set
/// per type is derived by
/// <see cref="EntityTypeCatalog.GetAllowedDateLabelPositions"/>.
/// </summary>
/// <remarks>
/// <para>
/// The set a type may name is derived from
/// <see cref="EntityTypeDefinition.DateMode"/>, never authored as a
/// second list: a dateless type names none (the control is disabled),
/// a <c>StartOnly</c> type names <c>Auto</c>, <c>Start</c>, and
/// <c>None</c>, and a span names all five. <c>Finish</c> and
/// <c>Both</c> are therefore never offered for a <c>StartOnly</c> type,
/// which reads no finish date.
/// </para>
/// </remarks>
public enum GanttDateLabelPosition
{
    /// <summary>
    /// The type default. Retained so a workbook written before the
    /// 2026-11-13 ruling — where <c>Auto</c> meant "both labels" —
    /// keeps parsing; it now resolves to <see cref="None"/>, the same
    /// as a blank cell, and the ribbon presents it as <c>Default</c>.
    /// </summary>
    Auto = 0,

    /// <summary>The start-date label only.</summary>
    Start = 1,

    /// <summary>The finish-date label only (span types only).</summary>
    Finish = 2,

    /// <summary>No date labels (the row's dates are hidden).</summary>
    None = 3,

    /// <summary>
    /// The start and the finish label together (span types only). It is an
    /// explicit choice, and the one a user picks to see both dates on a bar;
    /// the default is <see cref="None"/> (owner ruling 2026-11-13).
    /// </summary>
    Both = 4,
}
