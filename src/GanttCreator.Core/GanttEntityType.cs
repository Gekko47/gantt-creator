namespace GanttCreator.Core;

/// <summary>
/// A user-selectable row type in the <c>Type</c> column of the visible
/// <c>tblGanttData</c> table. The enum member names and numeric values are the
/// durable code identity. The exact workbook display names (which contain
/// spaces and hyphens and cannot be enum identifiers) are owned by
/// <see cref="EntityTypeCatalog"/> and form part of the workbook schema.
/// Renaming or localising a display name is a schema migration, but that
/// migration must not renumber or rename this machine identity (see
/// <see cref="GanttSchemaVersion"/>).
/// </summary>
public enum GanttEntityType
{
    /// <summary>Section header that splits the chart into named sections.</summary>
    Splitter = 0,

    /// <summary>Blank vertical space; renders no entity.</summary>
    Spacer = 1,

    /// <summary>As-built activity span.</summary>
    AsBuiltActivity = 2,

    /// <summary>As-planned activity span.</summary>
    AsPlannedActivity = 3,

    /// <summary>Baseline activity span.</summary>
    BaselineActivity = 4,

    /// <summary>Critical child interval span owned by a parent activity.</summary>
    CriticalInterval = 5,

    /// <summary>Explicit delay span.</summary>
    DelayEvent = 6,

    /// <summary>As-built procurement span.</summary>
    AsBuiltProcurement = 7,

    /// <summary>As-planned procurement span.</summary>
    AsPlannedProcurement = 8,

    /// <summary>Baseline procurement span.</summary>
    BaselineProcurement = 9,

    /// <summary>Named custom span; the row must supply a <c>StyleKey</c>.</summary>
    CustomActivity = 10,

    /// <summary>As-built milestone point event.</summary>
    AsBuiltMilestone = 11,

    /// <summary>As-planned milestone point event.</summary>
    AsPlannedMilestone = 12,

    /// <summary>Baseline milestone point event.</summary>
    BaselineMilestone = 13,

    /// <summary>Critical milestone point event.</summary>
    CriticalMilestone = 14,

    /// <summary>Full-height vertical date line; <c>Start</c> is its single date.</summary>
    Delineator = 15,
}
