using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R3.14 cross-renderer primitive-equivalence thin slice: one bar, one external
/// description label, and one milestone diamond are walked through the real built
/// scene and every field the entity guide's equivalence table requires of all four
/// renderers is asserted.
/// </summary>
/// <remarks>
/// <para>
/// This proves the <em>model</em>, not a render. The Excel live worksheet, the
/// editable composition, PowerPoint, and the PNG renderer all read the same scene
/// primitive, so a field the scene does not carry is a field one of them would have
/// to invent - which the guide forbids ("Renderers consume resolved scene
/// primitives ... must not independently move labels, recalculate dates, substitute
/// colours, change line widths, or reorder entities"). A missing field is therefore a
/// defect in the model, never a renderer-side workaround (R3.14 D1).
/// </para>
/// <para>
/// No renderer exists yet, so the proof that all four can read the same primitive is
/// that the field set survives <see cref="SceneSnapshot"/> canonical serialization
/// intact: the three round-trip facts below re-evaluate the identical catalogue
/// against the deserialized scene. If serialization alone could not prove a field,
/// R3.14 D2 requires a spike and an ADR instead of a model extension; no such gap
/// was found.
/// </para>
/// <para>
/// Each row's fields live in <see cref="EquivalenceFields"/>, transcribed from the
/// entity guide with a section citation, and the citation guard keeps that
/// transcription honest.
/// </para>
/// </remarks>
public sealed class EquivalenceThinSliceTests
{
    // The reference fixture's first As-Planned Activity row and its first milestone
    // row. Identifiers are fixed in tests/fixtures/reference-gantt.json, so the
    // primitives under test are stable and no row is generated here.
    private static readonly GanttRowId _barRow = GanttRowId.Parse("G-000000000000000000000000000000b1");
    private static readonly GanttRowId _milestoneRow = GanttRowId.Parse("G-000000000000000000000000000000f1");

    private static readonly SceneOwnerId _barOwner = SceneOwnerId.ForRow(_barRow);
    private static readonly SceneOwnerId _milestoneOwner = SceneOwnerId.ForRow(_milestoneRow);

    private static readonly string _barId = ScenePrimitive.CreateId(_barOwner, "bar");
    private static readonly string _labelId = ScenePrimitive.CreateId(_barOwner, "label");
    private static readonly string _markerId = ScenePrimitive.CreateId(_milestoneOwner, "marker");

    [Fact]
    public void The_span_bar_carries_every_field_the_equivalence_row_requires()
    {
        GanttScene scene = SceneBuilderTests.BuildScene();

        AssertRowIsSatisfied(EquivalenceFields.SpanBar, scene, _barId);
    }

    [Fact]
    public void The_external_description_label_carries_every_field_the_equivalence_row_requires()
    {
        GanttScene scene = SceneBuilderTests.BuildScene();

        AssertRowIsSatisfied(EquivalenceFields.ExternalDescriptionLabel, scene, _labelId);
    }

    [Fact]
    public void The_milestone_diamond_carries_every_field_the_equivalence_row_requires()
    {
        GanttScene scene = SceneBuilderTests.BuildScene();

        AssertRowIsSatisfied(EquivalenceFields.MilestoneDiamond, scene, _markerId);
    }

    [Theory]
    [InlineData("rect")]
    [InlineData("text")]
    [InlineData("polygon")]
    public void Every_required_field_survives_snapshot_serialization(string kind)
    {
        // The four renderers are separate code paths reading the same scene, so the
        // field set has to survive the canonical snapshot that all of them would be
        // handed or rebuilt from. Re-evaluating the identical catalogue against the
        // deserialized scene is that proof.
        GanttScene scene = SceneBuilderTests.BuildScene();
        GanttScene round = SceneSnapshot.Deserialize(SceneSnapshot.Serialize(scene));

        (EquivalenceRow row, string primitiveId) = kind switch
        {
            "rect" => (EquivalenceFields.SpanBar, _barId),
            "text" => (EquivalenceFields.ExternalDescriptionLabel, _labelId),
            _ => (EquivalenceFields.MilestoneDiamond, _markerId),
        };

        AssertRowIsSatisfied(row, round, primitiveId);
    }

    [Fact]
    public void A_bar_stripped_of_its_resolved_geometry_is_reported_as_missing_that_field()
    {
        // The positive test for the row: a field-stripped scene must fail the same
        // assertion, and the reported field must be the stripped one. Without this,
        // a collector that silently found nothing would pass every other test here.
        GanttScene stripped = ReplacePrimitive(SceneBuilderTests.BuildScene(), _barId, WithoutBarGeometry);

        Assert.Equal<string>(
            ["resolved bar geometry with non-zero extents"],
            EquivalenceFields.SpanBar.UnsatisfiedFields(stripped, _barId));
    }

    [Fact]
    public void A_label_stripped_of_its_resolved_text_bounds_is_reported_as_missing_that_field()
    {
        GanttScene stripped = ReplacePrimitive(SceneBuilderTests.BuildScene(), _labelId, WithoutTextBounds);

        Assert.Equal<string>(
            ["the resolved text bounds"],
            EquivalenceFields.ExternalDescriptionLabel.UnsatisfiedFields(stripped, _labelId));
    }

    [Fact]
    public void A_three_point_milestone_marker_is_reported_as_missing_the_diamond_geometry()
    {
        // A triangle is the realistic model defect here: a builder that emitted a
        // polygon of the wrong vertex count still produces a valid ScenePolygon, so
        // only this assertion can catch it. Dropping the left tip also halves one
        // axis, so the tip-to-tip field fails with it.
        GanttScene stripped = ReplacePrimitive(SceneBuilderTests.BuildScene(), _markerId, WithoutLeftTip);

        Assert.Equal<string>(
            ["the four points in draw order (top, right, bottom, left)", "the tip-to-tip extent, MilestoneSizePt on both axes"],
            EquivalenceFields.MilestoneDiamond.UnsatisfiedFields(stripped, _markerId));
    }

    [Fact]
    public void An_absent_primitive_is_reported_rather_than_passing_silently()
    {
        // A renamed or dropped role must not turn the row into a vacuous pass, so the
        // collector names the missing primitive instead of returning nothing.
        GanttScene scene = SceneBuilderTests.BuildScene();
        const string absentId = "G-000000000000000000000000000000b1:nonesuch";

        Assert.Equal<string>(
            [$"the primitive '{absentId}' is absent from the scene"],
            EquivalenceFields.ExternalDescriptionLabel.UnsatisfiedFields(scene, absentId));
    }

    [Fact]
    public void Every_transcribed_field_cites_the_entity_guide_and_is_named_once_per_row()
    {
        // The work item's main risk is the field list drifting from the guide. A
        // citation is what makes a drift visible in review, and a duplicate or blank
        // name would make a failure message ambiguous.
        foreach (EquivalenceRow row in EquivalenceFields.All)
        {
            Assert.NotEmpty(row.Fields);
            Assert.All(row.Fields, field =>
            {
                Assert.False(string.IsNullOrWhiteSpace(field.Name), row.Name + " has a blank field name.");
                Assert.Contains("Entity guide", field.GuideCitation, StringComparison.Ordinal);
            });

            Assert.Equal(
                row.Fields.Select(field => field.Name).Distinct(StringComparer.Ordinal).Count(),
                row.Fields.Count);
        }
    }

    private static void AssertRowIsSatisfied(EquivalenceRow row, GanttScene scene, string primitiveId)
    {
        IReadOnlyList<string> unsatisfied = row.UnsatisfiedFields(scene, primitiveId);
        Assert.True(
            unsatisfied.Count == 0,
            $"The {row.Name} does not carry every field the equivalence row requires. Missing or wrong: "
                + string.Join("; ", unsatisfied));
    }

    private static GanttScene ReplacePrimitive(GanttScene scene, string primitiveId, Func<ScenePrimitive, ScenePrimitive> replace)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentNullException.ThrowIfNull(replace);

        List<ScenePrimitive> primitives = [.. scene.Primitives.Select(primitive =>
            string.Equals(primitive.PrimitiveId, primitiveId, StringComparison.Ordinal) ? replace(primitive) : primitive)];

        SceneCreationOutcome outcome = GanttScene.TryCreate(scene.ChartBounds, scene.PlotBounds, primitives, scene.Warnings);
        Assert.True(outcome.Succeeded, "The field-stripped scene was refused: " + outcome.Refusal);
        return outcome.Scene!;
    }

    private static ScenePrimitive WithoutBarGeometry(ScenePrimitive primitive)
    {
        SceneRect bar = Assert.IsType<SceneRect>(primitive);
        return new SceneRect(
            bar.PrimitiveId,
            bar.OwnerId,
            bar.ZLayer,
            new RectD(0, 0, 0, 0),
            bar.Style,
            bar.EntityType,
            bar.LaneOrder,
            bar.StackIndex,
            bar.SortOrder);
    }

    private static ScenePrimitive WithoutTextBounds(ScenePrimitive primitive)
    {
        SceneText label = Assert.IsType<SceneText>(primitive);
        return new SceneText(
            label.PrimitiveId,
            label.OwnerId,
            label.ZLayer,
            label.Text,
            new RectD(0, 0, 0, 0),
            label.Style,
            label.Alignment,
            label.EntityType,
            label.LaneOrder,
            label.StackIndex,
            label.SortOrder);
    }

    private static ScenePrimitive WithoutLeftTip(ScenePrimitive primitive)
    {
        ScenePolygon diamond = Assert.IsType<ScenePolygon>(primitive);
        return new ScenePolygon(
            diamond.PrimitiveId,
            diamond.OwnerId,
            diamond.ZLayer,
            [.. diamond.Points.Take(3)],
            diamond.Style,
            diamond.EntityType,
            diamond.LaneOrder,
            diamond.StackIndex,
            diamond.SortOrder);
    }
}
