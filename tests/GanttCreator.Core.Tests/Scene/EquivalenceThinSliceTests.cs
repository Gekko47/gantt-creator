using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The cross-renderer primitive-equivalence proof: every field-contract row of the
/// entity guide is walked through a real built scene and asserted, and each row ships
/// a field-stripped positive proving its fields are load-bearing.
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
/// intact: the round-trip fact re-evaluates the identical catalogue against the
/// deserialized scene. If serialization alone could not prove a field, R3.14 D2
/// requires a spike and an ADR instead of a model extension; no such gap was found.
/// </para>
/// <para>
/// <b>The panel row needs its own scene.</b> The canonical reference build omits the
/// optional panel theme, so the committed golden contains no panel primitive. The
/// panel row is therefore asserted against <see cref="SceneBuilderTests.BuildSceneWithPanel"/>,
/// which adds only the panel and never changes the golden.
/// </para>
/// </remarks>
public sealed class EquivalenceThinSliceTests
{
    // The role-derived identifiers the field-stripped positives rebuild. Each is
    // derived from the fixture row id through the same ScenePrimitive.CreateId the
    // builders use, so a renamed role fails loudly here rather than silently addressing
    // a primitive that no longer exists.
    private static readonly string _barId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.BarProbeRow), "bar");
    private static readonly string _labelId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.BarProbeRow), "label");
    private static readonly string _markerId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.MilestoneProbeRow), "marker");
    private static readonly string _criticalId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.CriticalProbeRow), "critical");
    private static readonly string _sharedDelineatorId = SceneBuilderTests.SharedDelineatorPrimitiveId;
    private static readonly string _delineatorLabelId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.DelineatorProbeRow), "delineator-label");
    private static readonly string _panelBodyCellId = ScenePrimitive.CreateId(SceneOwnerId.ForRow(SceneBuilderTests.PanelProbeRow), "panel-cell:Id");
    private static readonly string _yearBandId = ScenePrimitive.CreateId(SceneOwnerId.Chart, "year:2026");
    private static readonly string _yearLabelId = ScenePrimitive.CreateId(SceneOwnerId.Chart, "year:2026:label");
    private static readonly string _backgroundId = ScenePrimitive.CreateId(SceneOwnerId.Chart, "background");

    private static GanttScene ReferenceScene() => SceneBuilderTests.BuildScene();

    /// <summary>
    /// Gets the rows the canonical reference scene can satisfy.
    /// </summary>
    /// <remarks>
    /// Every row except the data panel. The canonical build deliberately omits the
    /// optional panel theme, so the panel row can only be asserted against
    /// <see cref="SceneBuilderTests.BuildSceneWithPanel"/>. Filtering it here rather
    /// than skipping it is deliberate: the panel row is proven satisfiable by its own
    /// test, so nothing is left unproven.
    /// </remarks>
    private static IEnumerable<EquivalenceRow> RowsInReferenceScene() =>
        EquivalenceFields.All.Where(row => !row.IsExclusion && row != EquivalenceFields.DataPanel);

    [Fact]
    public void Every_row_of_the_field_contract_table_carries_every_field_it_requires()
    {
        GanttScene scene = ReferenceScene();

        // Every row, not a hand-picked subset: the catalogue's whole point is that a
        // renderer author can consult any row and find a proven field list. The panel
        // row has its own test below, against a panel-bearing build.
        foreach (EquivalenceRow row in RowsInReferenceScene())
        {
            AssertRowIsSatisfied(row, scene);
        }
    }

    [Fact]
    public void The_data_panel_row_carries_every_field_it_requires()
    {
        // The canonical scene carries no panel primitive, so this row is asserted
        // against a panel-bearing build instead. Proving the row is satisfiable at all
        // is the point: a row nobody can satisfy is documentation, not a contract.
        GanttScene scene = SceneBuilderTests.BuildSceneWithPanel();

        // Non-vacuity: the row must be evaluated against a scene that really has the
        // primitives it names, or every field would trivially "pass" as absent.
        Assert.NotNull(EquivalenceRow.FirstIdentifierContaining(scene, "panel-cell:"));
        Assert.NotNull(EquivalenceRow.FirstIdentifierContaining(scene, "header-cell:"));

        AssertRowIsSatisfied(EquivalenceFields.DataPanel, scene);
    }

    [Fact]
    public void The_excluded_rows_carry_no_primitive_a_renderer_could_invent()
    {
        // The validation indicator and the legend are rows the model deliberately does
        // not carry. Asserting their absence is what stops "a renderer must not invent
        // one" from being an unfalsifiable sentence in the guide.
        GanttScene scene = ReferenceScene();

        EquivalenceRow[] excluded = [.. EquivalenceFields.All.Where(row => row.IsExclusion)];
        Assert.NotEmpty(excluded);

        foreach (EquivalenceRow row in excluded)
        {
            string? offender = EquivalenceRow.FirstIdentifierContaining(scene, row.Exclusion!);
            Assert.True(
                offender is null,
                $"The {row.Name} row excludes '{row.Exclusion}', but the scene carries '{offender}'.");
        }
    }

    [Fact]
    public void Every_required_field_survives_snapshot_serialization()
    {
        // The four renderers are separate code paths reading the same scene, so the
        // field set has to survive the canonical snapshot that all of them would be
        // handed or rebuilt from. Re-evaluating the identical catalogue against the
        // deserialized scene is that proof.
        GanttScene round = SceneSnapshot.Deserialize(SceneSnapshot.Serialize(ReferenceScene()));

        foreach (EquivalenceRow row in RowsInReferenceScene())
        {
            AssertRowIsSatisfied(row, round);
        }
    }

    [Fact]
    public void A_bar_stripped_of_its_resolved_geometry_is_reported_as_missing_that_field()
    {
        // The positive test for the activity row: a field-stripped scene must fail the
        // same assertion, and the reported field must be the stripped one. Without
        // this, a collector that silently found nothing would pass every other test.
        GanttScene stripped = ReplacePrimitive(ReferenceScene(), _barId, WithoutBarGeometry);

        Assert.Equal<string>(
            ["the bar geometry with non-zero extents"],
            EquivalenceFields.ActivityBody.UnsatisfiedFields(stripped));
    }

    [Fact]
    public void A_label_stripped_of_its_resolved_text_bounds_is_reported_as_missing_that_field()
    {
        GanttScene stripped = ReplacePrimitive(ReferenceScene(), _labelId, WithoutTextBounds);

        Assert.Contains(
            "the description resolved text bounds",
            EquivalenceFields.DescriptionAndDateLabels.UnsatisfiedFields(stripped));
    }

    [Fact]
    public void A_three_point_milestone_marker_is_reported_as_missing_the_diamond_geometry()
    {
        // A triangle is the realistic model defect here: a builder that emitted a
        // polygon of the wrong vertex count still produces a valid ScenePolygon, so
        // only this assertion can catch it. Dropping the left tip also halves one
        // axis, so the tip-to-tip field fails with it.
        GanttScene stripped = ReplacePrimitive(ReferenceScene(), _markerId, WithoutLeftTip);

        Assert.Equal<string>(
            ["the four points in draw order (top, right, bottom, left)", "the tip-to-tip extent, MilestoneSizePt on both axes"],
            EquivalenceFields.MilestoneDiamond.UnsatisfiedFields(stripped));
    }

    [Fact]
    public void A_milestone_marker_whose_vertices_are_reordered_is_reported_as_missing_the_draw_order()
    {
        // Four points in the wrong order still form a valid ScenePolygon and still
        // satisfy the tip-to-tip extent, so only the order assertion can catch it: a
        // renderer walking the vertices in that order draws a different shape.
        GanttScene rotated = ReplacePrimitive(ReferenceScene(), _markerId, WithRotatedVertices);

        Assert.Equal<string>(
            ["the four points in draw order (top, right, bottom, left)"],
            EquivalenceFields.MilestoneDiamond.UnsatisfiedFields(rotated));
    }

    [Fact]
    public void An_absent_primitive_is_reported_rather_than_passing_silently()
    {
        // A renamed or dropped role must not turn the row into a vacuous pass, so the
        // collector names the missing primitive instead of returning nothing. The row
        // is narrowed to a single field so the assertion is about the reporting, not
        // about how many fields happen to fail.
        const string absentId = "G-000000000000000000000000000000b1:nonesuch";
        EquivalenceField kindField = EquivalenceFields.ActivityBody.Fields
            .Single(field => field.Name == "the bar primitive kind is rect");
        EquivalenceRow renamed = EquivalenceFields.ActivityBody with
        {
            Members = [new EquivalenceMember("bar", absentId)],
            Fields = [kindField],
        };

        Assert.Equal<string>(
            [$"the primitive '{absentId}' is absent from the scene"],
            renamed.UnsatisfiedFields(ReferenceScene()));
    }

    [Fact]
    public void A_critical_overlay_of_the_wrong_height_is_reported_as_missing_the_line_height()
    {
        // §11's overlay is CriticalLinePt tall and a renderer draws a line along its top
        // edge. An overlay that is a thick band instead still passes every other field
        // in the row, so only this assertion catches it.
        GanttScene fat = ReplacePrimitive(ReferenceScene(), _criticalId, WithThickCriticalOverlay);

        Assert.Equal<string>(
            ["the CriticalLinePt overlay height"],
            EquivalenceFields.CriticalInterval.UnsatisfiedFields(fat));
    }

    [Fact]
    public void A_delineator_line_that_stops_short_of_the_plot_is_reported_as_missing_the_full_height()
    {
        // §24 makes a delineator a full-plot-height line. A line that only spans part
        // of the plot is still a valid SceneLine with a resolved stroke, so every other
        // field in the row still passes.
        GanttScene shortened = ReplacePrimitive(ReferenceScene(), _sharedDelineatorId, WithShortenedDelineator);

        Assert.Equal<string>(
            ["the line spans the plot's full height"],
            EquivalenceFields.Delineator.UnsatisfiedFields(shortened));
    }

    [Fact]
    public void A_body_cell_fill_at_the_wrong_z_layer_is_reported_as_missing_that_layer()
    {
        // The R3.16 correction, exercised: §3 puts a body cell fill at Background, not
        // at Frame with the header. This is the positive for the corrected field - a
        // panel cell at the old layer fails it and only it.
        GanttScene scene = SceneBuilderTests.BuildSceneWithPanel();
        GanttScene wrong = ReplacePrimitive(scene, _panelBodyCellId, AtFrameLayer);

        Assert.Equal<string>(
            ["the body-cell z-layer"],
            EquivalenceFields.DataPanel.UnsatisfiedFields(wrong));
    }

    [Fact]
    public void A_year_label_whose_parent_rectangle_is_missing_is_reported_as_missing_that_field()
    {
        // R3.17's convention as a field: a year label names its parent by appending
        // ":label". Renaming the *band* leaves the label itself present and well-formed,
        // but orphaned - the realistic defect, because a renderer would then reconcile
        // the label against a rectangle that no longer exists. Renaming the label
        // instead would only prove the absent-primitive path, which is covered above.
        GanttScene orphaned = RenamePrimitive(ReferenceScene(), _yearBandId, "chart:year:2026-renamed");

        // The label must still be in the scene, or this would prove nothing about the
        // parent rule and only repeat the absent-primitive case.
        Assert.Contains(
            orphaned.Primitives,
            primitive => string.Equals(primitive.PrimitiveId, _yearLabelId, StringComparison.Ordinal));

        Assert.Contains(
            "the label names an existing parent rectangle by appending ':label'",
            EquivalenceFields.YearHeader.UnsatisfiedFields(orphaned));
    }

    [Fact]
    public void A_chart_background_that_disagrees_with_the_declared_chart_bounds_is_reported()
    {
        // R3.15's contract from the catalogue side. Resizing only the background leaves
        // the scene's declared ChartBounds untouched, which is exactly the two-sources
        // disagreement R3.15 removed from the builder.
        GanttScene grown = ReplacePrimitive(ReferenceScene(), _backgroundId, WithLargerBounds);

        Assert.Contains(
            "the scene ChartBounds equal the chart background",
            EquivalenceFields.BandsGridFrame.UnsatisfiedFields(grown));
    }

    [Fact]
    public void A_bar_re_owned_by_the_chart_is_reported_as_missing_its_row_ownership()
    {
        // Ownership drives reconciliation: a row-owned primitive dies with its row,
        // while chart furniture is never a deletion candidate. Re-owning a bar under
        // the chart is the realistic defect, and every other field in the row still
        // passes, so only the ownership field catches it.
        GanttScene chartOwned = ReplacePrimitive(ReferenceScene(), _barId, ReOwnedByChart);

        // Two fields fail, and both should: a row-owned primitive's identifier is
        // derived from its owner, so re-owning it also invalidates the role-derived
        // id. Asserting only one would hide half the defect.
        Assert.Equal<string>(
            ["the role-derived ':bar' identifier", "the bar owning row identity"],
            EquivalenceFields.ActivityBody.UnsatisfiedFields(chartOwned));
    }

    [Fact]
    public void A_delineator_label_given_lane_and_stack_keys_is_reported_as_missing_their_absence()
    {
        // The guide's Delineator row states the label carries no lane or stack keys -
        // it is positioned on the plot, not in a lane. Giving it keys is the realistic
        // defect, and a test asserting the keys are *present* (as the description-label
        // row does) would miss it entirely.
        GanttScene keyed = ReplacePrimitive(ReferenceScene(), _delineatorLabelId, WithLaneAndStackKeys);

        Assert.Equal<string>(
            ["the label carries no lane or stack order keys"],
            EquivalenceFields.Delineator.UnsatisfiedFields(keyed));
    }

    [Fact]
    public void Every_transcribed_field_cites_the_entity_guide_and_is_named_once_per_row()
    {
        // The work item's main risk is the field list drifting from the guide. A
        // citation is what makes a drift visible in review, a duplicate or blank name
        // would make a failure message ambiguous, and a field pointing at an
        // undeclared member would silently pass every scene.
        foreach (EquivalenceRow row in EquivalenceFields.All)
        {
            Assert.False(
                string.IsNullOrWhiteSpace(row.GuideRowName),
                "Every row must spell the guide's own row name for the drift guard.");

            if (row.IsExclusion)
            {
                // An exclusion row asserts an absence, so it carries no fields - but it
                // must still name the token it excludes, or it proves nothing.
                Assert.False(string.IsNullOrWhiteSpace(row.Exclusion), row.Name + " excludes nothing.");
                Assert.Empty(row.Fields);
                Assert.Empty(row.Members);
                continue;
            }

            Assert.NotEmpty(row.Fields);
            Assert.NotEmpty(row.Members);
            Assert.All(row.Fields, field =>
            {
                Assert.False(string.IsNullOrWhiteSpace(field.Name), row.Name + " has a blank field name.");
                Assert.Contains("Entity guide", field.GuideCitation, StringComparison.Ordinal);
                Assert.Contains(
                    row.Members,
                    member => string.Equals(member.Key, field.Member, StringComparison.Ordinal));
            });

            Assert.Equal(
                row.Fields.Select(field => field.Name).Distinct(StringComparer.Ordinal).Count(),
                row.Fields.Count);
        }
    }

    [Fact]
    public void Every_row_is_declared_exactly_once()
    {
        // All is what the drift guard and the all-rows fact iterate, so a row declared
        // twice would silently double its coverage while a row declared not at all would
        // never be asserted.
        Assert.Equal(
            EquivalenceFields.All.Count,
            EquivalenceFields.All.Select(row => row.GuideRowName).Distinct(StringComparer.Ordinal).Count());
    }

    private static void AssertRowIsSatisfied(EquivalenceRow row, GanttScene scene)
    {
        IReadOnlyList<string> unsatisfied = row.UnsatisfiedFields(scene);
        Assert.True(
            unsatisfied.Count == 0,
            $"The {row.Name} row does not carry every field the entity guide requires. Missing or wrong: "
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

    private static GanttScene RenamePrimitive(GanttScene scene, string primitiveId, string newId)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentException.ThrowIfNullOrWhiteSpace(newId);

        return ReplacePrimitive(
            scene,
            primitiveId,
            primitive => primitive switch
            {
                SceneRect rect => new SceneRect(newId, rect.OwnerId, rect.ZLayer, rect.Bounds, rect.Style, rect.EntityType, rect.LaneOrder, rect.StackIndex, rect.SortOrder),
                SceneText text => new SceneText(newId, text.OwnerId, text.ZLayer, text.Text, text.TextBounds, text.Style, text.Alignment, text.EntityType, text.LaneOrder, text.StackIndex, text.SortOrder),
                _ => throw new InvalidOperationException("RenamePrimitive supports only the kinds this test renames."),
            });
    }

    private static ScenePrimitive ReOwnedByChart(ScenePrimitive primitive)
    {
        SceneRect rect = Assert.IsType<SceneRect>(primitive);
        // The identifier is deliberately left naming the row while the owner now says
        // chart. That is the real reconciliation hazard - a renderer matching on the
        // shape name would still find the row's id - and keeping the id stable means
        // only the ownership field fails.
        return new SceneRect(
            rect.PrimitiveId,
            SceneOwnerId.Chart,
            rect.ZLayer,
            rect.Bounds,
            rect.Style,
            rect.EntityType,
            rect.LaneOrder,
            rect.StackIndex,
            rect.SortOrder);
    }

    private static ScenePrimitive WithLaneAndStackKeys(ScenePrimitive primitive)
    {
        SceneText text = Assert.IsType<SceneText>(primitive);
        return new SceneText(
            text.PrimitiveId,
            text.OwnerId,
            text.ZLayer,
            text.Text,
            text.TextBounds,
            text.Style,
            text.Alignment,
            text.EntityType,
            laneOrder: 0,
            stackIndex: 0,
            sortOrder: text.SortOrder);
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

    private static ScenePrimitive WithLargerBounds(ScenePrimitive primitive)
    {
        SceneRect rect = Assert.IsType<SceneRect>(primitive);
        return new SceneRect(
            rect.PrimitiveId,
            rect.OwnerId,
            rect.ZLayer,
            new RectD(rect.Bounds.X, rect.Bounds.Y, rect.Bounds.Width + 10, rect.Bounds.Height + 10),
            rect.Style,
            rect.EntityType,
            rect.LaneOrder,
            rect.StackIndex,
            rect.SortOrder);
    }

    private static ScenePrimitive AtFrameLayer(ScenePrimitive primitive)
    {
        SceneRect rect = Assert.IsType<SceneRect>(primitive);
        return new SceneRect(
            rect.PrimitiveId,
            rect.OwnerId,
            ZLayer.Frame,
            rect.Bounds,
            rect.Style,
            rect.EntityType,
            rect.LaneOrder,
            rect.StackIndex,
            rect.SortOrder);
    }

    private static ScenePrimitive WithThickCriticalOverlay(ScenePrimitive primitive)
    {
        SceneRect overlay = Assert.IsType<SceneRect>(primitive);
        return new SceneRect(
            overlay.PrimitiveId,
            overlay.OwnerId,
            overlay.ZLayer,
            new RectD(overlay.Bounds.X, overlay.Bounds.Y, overlay.Bounds.Width, overlay.Bounds.Height + 5),
            overlay.Style,
            overlay.EntityType,
            overlay.LaneOrder,
            overlay.StackIndex,
            overlay.SortOrder);
    }

    private static ScenePrimitive WithShortenedDelineator(ScenePrimitive primitive)
    {
        SceneLine line = Assert.IsType<SceneLine>(primitive);
        return new SceneLine(
            line.PrimitiveId,
            line.OwnerId,
            line.ZLayer,
            new PointD(line.From.X, line.From.Y + 5),
            new PointD(line.To.X, line.To.Y - 5),
            line.Style,
            line.EntityType,
            line.LaneOrder,
            line.StackIndex,
            line.SortOrder);
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

    private static ScenePrimitive WithRotatedVertices(ScenePrimitive primitive)
    {
        ScenePolygon diamond = Assert.IsType<ScenePolygon>(primitive);

        // One step round: right, bottom, left, top. Same four points, same bounds.
        return new ScenePolygon(
            diamond.PrimitiveId,
            diamond.OwnerId,
            diamond.ZLayer,
            [diamond.Points[1], diamond.Points[2], diamond.Points[3], diamond.Points[0]],
            diamond.Style,
            diamond.EntityType,
            diamond.LaneOrder,
            diamond.StackIndex,
            diamond.SortOrder);
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
