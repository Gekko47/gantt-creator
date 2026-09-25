using System.Text.Json;
using System.Text.Json.Nodes;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

public sealed class SceneSnapshotTests
{
    [Fact]
    public void Chart_owned_warning_round_trips_through_the_snapshot()
    {
        GanttScene scene = CreateScene([
            new SceneRect("chart:frame", SceneOwnerId.Chart, ZLayer.Frame, new RectD(0, 0, 10, 10), new SceneStyle("Frame")),
        ]);
        GanttScene withWarning = GanttScene
            .TryCreate(
                scene.ChartBounds,
                scene.PlotBounds,
                scene.Primitives,
                [new SceneWarning(SceneOwnerId.Chart, "TitleOverflow", "Title truncated.")]
            )
            .Scene!;

        var json = SceneSnapshot.Serialize(withWarning);
        GanttScene roundTripped = SceneSnapshot.Deserialize(json);

        Assert.Equal(SceneOwnerId.Chart, roundTripped.Primitives.Single().OwnerId);
        Assert.Equal(SceneOwnerId.Chart, roundTripped.Warnings.Single().OwnerId);
    }

    [Fact]
    public void Deserialize_rejects_unknown_owner_fields_and_values_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene());

        JsonNode document = JsonNode.Parse(json)!;
        document["Primitives"]![0]!["Owner"]!["Value"] = "chart";

        _ = Assert.Throws<InvalidDataException>(() =>
            SceneSnapshot.Deserialize(json.Replace("\"Kind\":\"row\"", "\"Kind\":\"other\"", StringComparison.Ordinal))
        );
        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(document.ToJsonString()));
    }

    [Fact]
    public void Serialize_deserialize_round_trips_to_the_same_canonical_json()
    {
        var scene = CreateScene();
        string json = SceneSnapshot.Serialize(scene);
        GanttScene roundTripped = SceneSnapshot.Deserialize(json);

        Assert.Equal(json, SceneSnapshot.Serialize(roundTripped));
    }

    [Fact]
    public void Serialize_is_byte_identical_for_three_shuffled_inputs()
    {
        var owner = SceneOwnerId.ForRow(GanttRowId.New());
        var first = new SceneRect(
            "row:planned",
            owner,
            ZLayer.ActivityBody,
            new RectD(1, 2, 3, 4),
            new SceneStyle("Planned"),
            GanttEntityType.AsPlannedActivity,
            1,
            0,
            1
        );
        var second = new SceneRect(
            "row:actual",
            owner,
            ZLayer.ActivityBody,
            new RectD(5, 6, 7, 8),
            new SceneStyle("Actual"),
            GanttEntityType.AsBuiltActivity,
            1,
            0,
            1
        );
        var line = new SceneLine("row:line", owner, ZLayer.Grid, new PointD(0, 0), new PointD(10, 10), new SceneStyle("Grid"));
        var polygon = new ScenePolygon(
            "row:diamond",
            owner,
            ZLayer.Milestone,
            [new(0, 0), new(1, 0), new(0, 1)],
            new SceneStyle("Milestone"),
            GanttEntityType.AsBuiltMilestone
        );
        var text = new SceneText(
            "row:label",
            owner,
            ZLayer.Label,
            "Activity",
            new RectD(0, 0, 20, 10),
            new SceneStyle("Text"),
            GanttLabelPosition.Auto
        );
        var group = new SceneGroup("row:group", owner, ZLayer.Label, ["row:label", "row:diamond"]);
        ScenePrimitive[] input = [first, second, line, polygon, text, group];
        var expected = SceneSnapshot.Serialize(CreateScene(input));

        string[] shuffled =
        [
            SceneSnapshot.Serialize(CreateScene([group, text, polygon, line, second, first])),
            SceneSnapshot.Serialize(CreateScene([line, first, group, polygon, text, second])),
            SceneSnapshot.Serialize(CreateScene([text, second, group, first, polygon, line])),
        ];

        Assert.All(shuffled, actual => Assert.Equal(expected, actual));
    }

    [Fact]
    public void Deserialize_rejects_unknown_fields_as_a_json_error()
    {
        var json = SceneSnapshot
            .Serialize(CreateScene())
            .Replace("\"Version\":2", "\"Version\":2,\"Unknown\":true", StringComparison.Ordinal);

        _ = Assert.Throws<JsonException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_malformed_json_as_a_json_error() => Assert.Throws<JsonException>(() => SceneSnapshot.Deserialize("{"));

    [Fact]
    public void Deserialize_rejects_non_finite_numbers_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene()).Replace("\"X\":0", "\"X\":1e999", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_negative_extents_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene()).Replace("\"Width\":100", "\"Width\":-1", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_unknown_primitive_kind_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene()).Replace("\"Kind\":\"rect\"", "\"Kind\":\"unknown\"", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_invalid_colour_as_invalid_scene_data()
    {
        var json = SceneSnapshot
            .Serialize(CreateScene())
            .Replace("\"FillColour\":null", "\"FillColour\":\"invalid\"", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_undefined_hatch_pattern_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene()).Replace("\"HatchPattern\":0", "\"HatchPattern\":999", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_rejects_undefined_style_alignment_as_invalid_scene_data()
    {
        var json = SceneSnapshot.Serialize(CreateScene()).Replace("\"Alignment\":null", "\"Alignment\":999", StringComparison.Ordinal);

        _ = Assert.Throws<InvalidDataException>(() => SceneSnapshot.Deserialize(json));
    }

    [Fact]
    public void Deserialize_preserves_null_style_alignment()
    {
        GanttScene scene = SceneSnapshot.Deserialize(SceneSnapshot.Serialize(CreateScene()));

        Assert.Null(scene.Primitives.OfType<SceneRect>().Single().Style.Alignment);
    }

    private static GanttScene CreateScene() =>
        CreateScene([
            new SceneRect(
                "row:bar",
                SceneOwnerId.ForRow(GanttRowId.New()),
                ZLayer.ActivityBody,
                new RectD(1, 2, 3, 4),
                new SceneStyle("Planned"),
                GanttEntityType.AsPlannedActivity
            ),
        ]);

    private static GanttScene CreateScene(IReadOnlyList<ScenePrimitive> primitives)
    {
        var outcome =
            GanttScene.TryCreate(new RectD(0, 0, 100, 100), new RectD(10, 10, 90, 90), primitives, []).Scene
            ?? throw new InvalidOperationException("Fixture scene creation failed.");
        return outcome;
    }
}
