namespace GanttCreator.Core.Scene;

/// <summary>The reason an explicit time scale could not be created.</summary>
public enum TimeScaleCreationRefusal
{
    /// <summary>The plot start is after the plot finish.</summary>
    StartAfterFinish = 0,

    /// <summary>Either plot date is the default <see cref="DateOnly"/> value.</summary>
    DefaultDates = 1,

    /// <summary>The plot geometry is non-finite or has a non-positive width.</summary>
    NonFiniteOrDegenerateWidth = 2,
}

/// <summary>The typed result of attempting to create a <see cref="TimeScale"/>.</summary>
/// <param name="Scale">The created time scale when successful.</param>
/// <param name="Refusal">The typed refusal when unsuccessful.</param>
public sealed record TimeScaleCreationOutcome(
    TimeScale? Scale,
    TimeScaleCreationRefusal? Refusal)
{
    /// <summary>Gets a value indicating whether creation succeeded.</summary>
    public bool Succeeded => Scale is not null;
}
