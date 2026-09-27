namespace GanttCreator.Core.Scene;

/// <summary>One milestone marker request.</summary>
/// <param name="Event">The validated milestone event. Only <c>Start</c> is read.</param>
/// <param name="Style">The resolved scene style for the diamond.</param>
/// <param name="MilestoneSizePt">The tip-to-tip width and height, in points.</param>
/// <param name="SlotCentreY">The vertical centre of the visual stack slot, in points.</param>
/// <param name="LaneOrder">The lane ordering value, when known.</param>
/// <param name="StackIndex">The stack ordering value, when known.</param>
public sealed record MilestoneMarkerRequest(
    GanttEvent Event,
    SceneStyle Style,
    double MilestoneSizePt,
    double SlotCentreY,
    int? LaneOrder = null,
    int? StackIndex = null
);

/// <summary>The result of building one milestone marker.</summary>
/// <param name="Primitive">The diamond, or <see langword="null"/> when the date lies wholly outside the plot.</param>
/// <param name="CentreX">The exact centre X, or <see langword="null"/> when no marker was emitted.</param>
/// <param name="Warnings">The deterministic non-blocking warnings for this marker.</param>
public sealed record MilestoneMarkerResult(
    ScenePolygon? Primitive,
    double? CentreX,
    IReadOnlyList<SceneWarning> Warnings
);

/// <summary>The reason a milestone marker could not be built.</summary>
public enum MilestoneMarkerRefusal
{
    /// <summary>The request was null.</summary>
    NullRequest = 0,

    /// <summary>The event or its style was null.</summary>
    NullDependency = 1,

    /// <summary>The time scale was invalid.</summary>
    InvalidTimeScale = 2,

    /// <summary>The size, slot centre, or an edge was not finite, or the size was not positive.</summary>
    InvalidGeometry = 3,

    /// <summary>The event carries no start date, so a point event has nothing to place.</summary>
    MissingEventDate = 4,

    /// <summary>The event is not a milestone type.</summary>
    NotAMilestone = 5,
}

/// <summary>The typed result of attempting to build a milestone marker.</summary>
/// <param name="Result">The successful result, or <see langword="null"/>.</param>
/// <param name="Refusal">The refusal reason, or <see langword="null"/>.</param>
public sealed record MilestoneMarkerCreationOutcome(MilestoneMarkerResult? Result, MilestoneMarkerRefusal? Refusal)
{
    /// <summary>Gets whether construction succeeded.</summary>
    public bool Succeeded => Result is not null;
}

/// <summary>
/// Builds deterministic milestone-diamond geometry, with no Office dependency
/// and no text measurement.
/// </summary>
public static class MilestoneMarkerBuilder
{
    /// <summary>The warning code emitted when a milestone date lies outside the plot.</summary>
    public const string OutsidePlotRangeCode = "MilestoneOutsidePlotRange";

    /// <summary>Attempts to build one milestone marker.</summary>
    /// <param name="request">The typed milestone marker request.</param>
    /// <param name="timeScale">The validated time scale.</param>
    /// <returns>A typed result or refusal.</returns>
    public static MilestoneMarkerCreationOutcome TryBuild(
        MilestoneMarkerRequest? request,
        TimeScale? timeScale
    )
    {
        if (request is null)
        {
            return Refused(MilestoneMarkerRefusal.NullRequest);
        }

        if (timeScale is null)
        {
            return Refused(MilestoneMarkerRefusal.InvalidTimeScale);
        }

        if (request.Event is null || request.Style is null)
        {
            return Refused(MilestoneMarkerRefusal.NullDependency);
        }

        if (
            !double.IsFinite(request.MilestoneSizePt)
            || request.MilestoneSizePt <= 0
            || !double.IsFinite(request.SlotCentreY)
        )
        {
            return Refused(MilestoneMarkerRefusal.InvalidGeometry);
        }

        GanttEvent @event = request.Event;
        EntityTypeDefinition? definition = EntityTypeCatalog.GetDefinition(@event.Type);
        if (definition is null || definition.Kind != EntityKind.Milestone)
        {
            return Refused(MilestoneMarkerRefusal.NotAMilestone);
        }

        // Entity guide §20: a milestone is a point event that reads Start only.
        // Finish, LaneId, and StackIndex are never consulted for geometry, and no
        // half-day is added or inherited from the activity duration rule.
        if (@event.Start is not { } date)
        {
            return Refused(MilestoneMarkerRefusal.MissingEventDate);
        }

        List<SceneWarning> warnings = [];

        // The `Try*` form is used deliberately: DateToX throws for a date outside
        // the scale, and an out-of-range milestone must produce a warning rather
        // than an exception. The failure sentinel (x = 0) is never inspected.
        if (!timeScale.TryDateToX(date, out var centreX))
        {
            warnings.Add(
                new SceneWarning(
                    SceneOwnerId.ForRow(@event.Id),
                    OutsidePlotRangeCode,
                    "The milestone date lies outside the plot range, so no marker was drawn.")
            );
            return new MilestoneMarkerCreationOutcome(new MilestoneMarkerResult(null, null, warnings), null);
        }

        // §20: a four-point polygon — not a rotated square — centred on the date X
        // and the slot-band centre Y, with MilestoneSizePt as the tip-to-tip
        // extent on BOTH axes, so the bounds are exact and consistent across
        // renderers.
        var half = request.MilestoneSizePt / 2;
        var centreY = request.SlotCentreY;
        PointD[] points =
        [
            new(centreX, centreY - half),
            new(centreX + half, centreY),
            new(centreX, centreY + half),
            new(centreX - half, centreY),
        ];

        var polygon = new ScenePolygon(
            $"{@event.Id.Value}:marker",
            SceneOwnerId.ForRow(@event.Id),
            ZLayer.Milestone,
            points,
            request.Style,
            @event.Type,
            request.LaneOrder,
            request.StackIndex,
            @event.SortOrder
        );

        return new MilestoneMarkerCreationOutcome(new MilestoneMarkerResult(polygon, centreX, warnings), null);
    }

    private static MilestoneMarkerCreationOutcome Refused(MilestoneMarkerRefusal refusal) => new(null, refusal);
}
