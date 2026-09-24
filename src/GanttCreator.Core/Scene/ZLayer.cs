namespace GanttCreator.Core.Scene;

/// <summary>The exact scene z-layers defined by the Gantt entity guide.</summary>
public enum ZLayer
{
    /// <summary>Chart and data-panel background.</summary>
    Background = 0,

    /// <summary>Alternating time bands.</summary>
    AlternateBand = 10,

    /// <summary>Grid lines and lane separators.</summary>
    Grid = 20,

    /// <summary>Delineator lines.</summary>
    Delineator = 25,

    /// <summary>Splitter and section backgrounds.</summary>
    Section = 30,

    /// <summary>Activity bodies.</summary>
    ActivityBody = 40,

    /// <summary>Critical interval overlays.</summary>
    CriticalOverlay = 50,

    /// <summary>Milestone diamonds.</summary>
    Milestone = 60,

    /// <summary>Activity, milestone, and date labels.</summary>
    Label = 70,

    /// <summary>Delineator labels.</summary>
    DelineatorLabel = 75,

    /// <summary>Plot frame, table borders, and headers.</summary>
    Frame = 80,

    /// <summary>Chart title and optional legend.</summary>
    Title = 90,
}
