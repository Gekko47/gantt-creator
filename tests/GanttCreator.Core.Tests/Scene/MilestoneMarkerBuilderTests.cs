using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class MilestoneMarkerBuilderTests
{
    // A 31-day January 2024 plot, 0..310pt, so one day is exactly 10pt.
    private static readonly TimeScale _scale =
        TimeScale.TryCreate(new DateOnly(2024, 1, 1), new DateOnly(2024, 1, 31), 0, 310).Scale!;

    private static readonly SceneStyle _style = new("PlannedMilestone", fillColour: ColourHex.Parse("#92D050"));

    [Fact]
    public void Builds_a_four_point_diamond_whose_tips_are_exactly_tip_to_tip()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 5));

        Assert.True(outcome.Succeeded);
        ScenePolygon diamond = Assert.IsType<ScenePolygon>(outcome.Result!.Primitive);

        // 5 Jan is the fifth day, so centre X is 40. With an 8pt tip-to-tip size
        // the half-extent is 4, and a slot centre of 60 puts the tips at 56/64.
        PointD[] expected =
        [
            new(40, 56),
            new(44, 60),
            new(40, 64),
            new(36, 60),
        ];
        Assert.Equal(expected, diamond.Points.ToArray());
    }

    [Fact]
    public void Uses_the_same_tip_to_tip_extent_on_both_axes()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 5), size: 12);

        ScenePolygon diamond = Assert.IsType<ScenePolygon>(outcome.Result!.Primitive);
        var left = diamond.Points.Min(point => point.X);
        var right = diamond.Points.Max(point => point.X);
        var top = diamond.Points.Min(point => point.Y);
        var bottom = diamond.Points.Max(point => point.Y);

        Assert.Equal(12, right - left, 10);
        Assert.Equal(12, bottom - top, 10);
    }

    [Fact]
    public void Emits_the_diamond_at_the_milestone_layer_with_a_role_derived_id()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 5));

        ScenePolygon diamond = Assert.IsType<ScenePolygon>(outcome.Result!.Primitive);
        Assert.Equal(ZLayer.Milestone, diamond.ZLayer);
        Assert.EndsWith(":marker", diamond.PrimitiveId, StringComparison.Ordinal);
    }

    [Fact]
    public void Maps_a_date_to_the_start_of_that_day_with_no_extra_day_or_half_day()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 5));

        // §20: the centre is exactly DateToX(Start). A point event inherits
        // neither the activity right edge nor a half-day offset.
        Assert.Equal(40, outcome.Result!.CentreX);
    }

    [Fact]
    public void Ignores_the_finish_lane_and_stack_fields_for_geometry()
    {
        // §20/checklist C: only Start is read. Populating the other fields must
        // not move the marker; R2.5 already warns that they are not used.
        MilestoneMarkerCreationOutcome outcome = Build(
            new DateOnly(2024, 1, 5),
            finish: new DateOnly(2024, 1, 20),
            laneId: GanttRowId.New(),
            stackIndex: 4
        );

        Assert.Equal(40, outcome.Result!.CentreX);
        ScenePolygon diamond = Assert.IsType<ScenePolygon>(outcome.Result.Primitive);
        Assert.Equal(40, diamond.Points[0].X);
    }

    [Fact]
    public void Renders_at_the_exact_left_plot_edge_for_the_first_date_in_range()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 1));

        Assert.True(outcome.Succeeded);
        Assert.Equal(0, outcome.Result!.CentreX);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void Renders_at_the_exact_right_plot_edge_for_the_last_date_in_range()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 31));

        Assert.True(outcome.Succeeded);
        Assert.Equal(300, outcome.Result!.CentreX);
        Assert.Empty(outcome.Result.Warnings);
    }

    [Fact]
    public void Emits_no_marker_and_one_warning_for_a_date_after_the_plot()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 2, 5));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Null(outcome.Result.CentreX);
        Assert.Equal(MilestoneMarkerBuilder.OutsidePlotRangeCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Emits_no_marker_and_one_warning_for_a_date_before_the_plot()
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2023, 12, 5));

        Assert.True(outcome.Succeeded);
        Assert.Null(outcome.Result!.Primitive);
        Assert.Equal(MilestoneMarkerBuilder.OutsidePlotRangeCode, Assert.Single(outcome.Result.Warnings).Code);
    }

    [Fact]
    public void Same_date_milestones_share_one_centre_x_deterministically()
    {
        DateOnly date = new(2024, 1, 10);
        MilestoneMarkerCreationOutcome first = Build(date, type: GanttEntityType.AsBuiltMilestone);
        MilestoneMarkerCreationOutcome second = Build(date, type: GanttEntityType.CriticalMilestone);

        // §9: same stack, same date, one centre, so the two markers coincide
        // exactly rather than one nudging the other. 10 Jan is the tenth day,
        // so the shared centre X is 90pt.
        Assert.Equal(first.Result!.CentreX, second.Result!.CentreX);
        Assert.Equal(90, first.Result.CentreX);
        Assert.Equal(90, second.Result.CentreX);
    }

    [Theory]
    [InlineData(GanttEntityType.AsBuiltMilestone)]
    [InlineData(GanttEntityType.AsPlannedMilestone)]
    [InlineData(GanttEntityType.BaselineMilestone)]
    [InlineData(GanttEntityType.CriticalMilestone)]
    public void Emits_the_supplied_style_for_every_milestone_subtype(GanttEntityType type)
    {
        MilestoneMarkerCreationOutcome outcome = Build(new DateOnly(2024, 1, 5), type: type);

        ScenePolygon diamond = Assert.IsType<ScenePolygon>(outcome.Result!.Primitive);
        Assert.Equal(_style, diamond.Style);
        Assert.Equal(type, diamond.EntityType);
    }

    [Fact]
    public void Produces_identical_geometry_across_repeated_runs()
    {
        PointD[] first = Assert.IsType<ScenePolygon>(Build(new DateOnly(2024, 1, 5)).Result!.Primitive).Points.ToArray();
        PointD[] second = Assert.IsType<ScenePolygon>(Build(new DateOnly(2024, 1, 5)).Result!.Primitive).Points.ToArray();
        PointD[] third = Assert.IsType<ScenePolygon>(Build(new DateOnly(2024, 1, 5)).Result!.Primitive).Points.ToArray();

        Assert.Equal(first, second);
        Assert.Equal(first, third);
    }

    [Fact]
    public void Refuses_a_null_request()
    {
        MilestoneMarkerCreationOutcome outcome = MilestoneMarkerBuilder.TryBuild(null, _scale);

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.NullRequest, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_null_time_scale()
    {
        MilestoneMarkerCreationOutcome outcome = MilestoneMarkerBuilder.TryBuild(
            new MilestoneMarkerRequest(Event(new DateOnly(2024, 1, 5)), _style, 8, 60),
            null
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.InvalidTimeScale, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_non_finite_milestone_size()
    {
        MilestoneMarkerCreationOutcome outcome = MilestoneMarkerBuilder.TryBuild(
            new MilestoneMarkerRequest(Event(new DateOnly(2024, 1, 5)), _style, double.NaN, 60),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_zero_milestone_size_that_would_collapse_the_diamond()
    {
        MilestoneMarkerCreationOutcome outcome = MilestoneMarkerBuilder.TryBuild(
            new MilestoneMarkerRequest(Event(new DateOnly(2024, 1, 5)), _style, 0, 60),
            _scale
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.InvalidGeometry, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_milestone_with_no_start_date()
    {
        MilestoneMarkerCreationOutcome outcome = Build(start: null);

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.MissingEventDate, outcome.Refusal);
    }

    [Fact]
    public void Refuses_a_non_milestone_type_rather_than_rendering_a_diamond()
    {
        MilestoneMarkerCreationOutcome outcome = Build(
            new DateOnly(2024, 1, 5),
            type: GanttEntityType.AsPlannedActivity
        );

        Assert.False(outcome.Succeeded);
        Assert.Equal(MilestoneMarkerRefusal.NotAMilestone, outcome.Refusal);
    }

    private static MilestoneMarkerCreationOutcome Build(
        DateOnly? start,
        DateOnly? finish = null,
        double size = 8,
        GanttEntityType type = GanttEntityType.AsPlannedMilestone,
        GanttRowId? laneId = null,
        int? stackIndex = null
    ) =>
        MilestoneMarkerBuilder.TryBuild(
            new MilestoneMarkerRequest(Event(start, finish, type, laneId, stackIndex), _style, size, 60),
            _scale
        );

    private static GanttEvent Event(
        DateOnly? start,
        DateOnly? finish = null,
        GanttEntityType type = GanttEntityType.AsPlannedMilestone,
        GanttRowId? laneId = null,
        int? stackIndex = null
    ) =>
        new(
            1,
            GanttRowId.New(),
            laneId,
            stackIndex,
            type,
            "Milestone",
            start,
            finish,
            null,
            null,
            null,
            null,
            null,
            true,
            null
        );
}
