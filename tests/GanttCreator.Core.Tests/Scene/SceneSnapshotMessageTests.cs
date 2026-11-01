using System.Text.Json.Nodes;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Exact-message pins for <see cref="SceneSnapshot"/>'s rejection paths.
/// </summary>
/// <remarks>
/// <para>
/// Mutation testing left 67 mutants open in `SceneSnapshot`, 51 of them string
/// mutants. A serialise/deserialise round-trip cannot kill the dominant class:
/// the field-name literals passed to <c>Required</c> are only read when a field
/// is <em>missing</em>, so a successful round-trip never touches them.
/// </para>
/// <para>
/// These fixtures therefore corrupt a document that the serializer itself
/// produced, using <see cref="JsonNode"/> surgery rather than hand-written JSON.
/// That keeps every fixture schema-accurate by construction, which matters
/// because the reader uses <c>UnmappedMemberHandling.Disallow</c> and would
/// reject an invented property before reaching the code under test.
/// </para>
/// </remarks>
public class SceneSnapshotMessageTests
{
    private static GanttScene EmptyScene()
    {
        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 100),
            new RectD(0, 0, 50, 50),
            [],
            []);
        Assert.NotNull(outcome.Scene);
        return outcome.Scene;
    }

    /// <summary>The canonical JSON of an empty scene, as the serializer emits it.</summary>
    private static string CanonicalJson() => SceneSnapshot.Serialize(EmptyScene());

    /// <summary>Removes one property from the document, leaving the rest canonical.</summary>
    private static string Without(string property)
    {
        JsonNode node = JsonNode.Parse(CanonicalJson())!;
        Assert.True(((JsonObject)node).Remove(property), $"fixture assumes the document carries '{property}'");
        return node.ToJsonString();
    }

    [Fact]
    public void A_scene_round_trips_through_canonical_json_unchanged()
    {
        // Deterministic and idempotent: the second serialisation must be
        // byte-identical to the first, or the snapshot is not canonical.
        string first = CanonicalJson();
        string second = SceneSnapshot.Serialize(SceneSnapshot.Deserialize(first));

        Assert.Equal(first, second);
    }

    [Fact]
    public void An_unsupported_snapshot_version_names_the_version()
    {
        JsonNode node = JsonNode.Parse(CanonicalJson())!;
        node["Version"] = 2;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(node.ToJsonString()));

        Assert.Equal("Unsupported scene snapshot version 2.", exception.Message);
    }

    [Fact]
    public void A_missing_chart_bounds_names_the_field()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Without("ChartBounds")));

        Assert.Equal("Missing scene field 'chartBounds'.", exception.Message);
    }

    [Fact]
    public void A_missing_plot_bounds_names_the_field()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Without("PlotBounds")));

        Assert.Equal("Missing scene field 'plotBounds'.", exception.Message);
    }

    /// <summary>Replaces the warning list with one warning object, then removes one property.</summary>
    private static string WarningWithout(params string[] removed)
    {
        JsonNode node = JsonNode.Parse(CanonicalJson())!;
        JsonObject warning = new()
        {
            ["Owner"] = new JsonObject { ["Kind"] = "chart", ["Value"] = "chart" },
            ["Code"] = "scene.test",
            ["Message"] = "test warning",
        };
        foreach (string property in removed)
        {
            Assert.True(warning.Remove(property), $"fixture assumes the warning carries '{property}'");
        }

        node["Warnings"] = new JsonArray(warning);
        return node.ToJsonString();
    }

    [Fact]
    public void A_warning_without_an_owner_names_the_field()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(WarningWithout("Owner")));

        Assert.Equal("Missing scene field 'warning.owner'.", exception.Message);
    }

    [Fact]
    public void A_warning_without_a_code_names_the_field()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(WarningWithout("Code")));

        Assert.Equal("Missing scene field 'warning.code'.", exception.Message);
    }

    [Fact]
    public void A_warning_without_a_message_names_the_field()
    {
        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(WarningWithout("Message")));

        Assert.Equal("Missing scene field 'warning.message'.", exception.Message);
    }

    /// <summary>A scene carrying one primitive of every kind, plus both owner kinds.</summary>
    private static GanttScene FullScene()
    {
        SceneStyle style = new("TestStyle", fillColour: null);
        SceneOwnerId row = SceneOwnerId.ForRow(RowId());
        ScenePrimitive[] primitives =
        [
            new SceneRect("p-rect", row, ZLayer.ActivityBody, new RectD(1, 2, 3, 4), style),
            new SceneLine("p-line", row, ZLayer.Grid, new PointD(1, 2), new PointD(3, 4), style),
            new ScenePolygon(
                "p-poly",
                row,
                ZLayer.CriticalOverlay,
                [new PointD(1, 2), new PointD(3, 4), new PointD(5, 6)],
                style),
            new SceneText(
                "p-text",
                row,
                ZLayer.Label,
                "hello",
                new RectD(1, 2, 3, 4),
                style,
                GanttTextAlignment.Centre),
            new SceneGroup("p-group", SceneOwnerId.Chart, ZLayer.Frame, ["p-rect"]),
        ];

        SceneCreationOutcome outcome = GanttScene.TryCreate(
            new RectD(0, 0, 100, 100),
            new RectD(0, 0, 50, 50),
            primitives,
            []);
        Assert.NotNull(outcome.Scene);
        return outcome.Scene;
    }

    private static GanttRowId RowId()
    {
        Assert.True(GanttRowId.TryParse($"G-{new string('7', 32)}", out GanttRowId? parsed));
        return parsed!;
    }

    private static JsonObject FullDocument() => (JsonObject)JsonNode.Parse(SceneSnapshot.Serialize(FullScene()))!;

    /// <summary>The single serialized primitive of the requested kind.</summary>
    private static JsonObject Primitive(JsonObject root, string kind) =>
        ((JsonArray)root["Primitives"]!)
            .Cast<JsonObject>()
            .Single(primitive => (string?)primitive["Kind"] == kind);

    private static string Serialized(JsonObject root) => root.ToJsonString();

    [Fact]
    public void A_scene_with_every_primitive_kind_round_trips_unchanged()
    {
        // Exercises the five primitive Kind literals and the chart/row owner
        // kinds: a drifted literal serialises to something that cannot be read
        // back, so byte equality of the second serialisation catches it.
        string first = SceneSnapshot.Serialize(FullScene());
        string second = SceneSnapshot.Serialize(SceneSnapshot.Deserialize(first));

        Assert.Equal(first, second);
    }

    [Fact]
    public void A_primitive_without_bounds_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "rect").Remove("Bounds"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'bounds'.", exception.Message);
    }

    [Fact]
    public void A_primitive_without_a_style_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "rect").Remove("Style"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'style'.", exception.Message);
    }

    [Fact]
    public void A_primitive_without_an_id_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "rect").Remove("PrimitiveId"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'primitiveId'.", exception.Message);
    }

    [Fact]
    public void A_primitive_without_an_owner_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "rect").Remove("Owner"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'primitive.owner'.", exception.Message);
    }

    [Fact]
    public void An_undefined_z_layer_is_named_with_its_value()
    {
        JsonObject root = FullDocument();
        Primitive(root, "rect")["ZLayer"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid z-layer 99.", exception.Message);
    }

    [Fact]
    public void A_style_without_a_key_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(((JsonObject)Primitive(root, "rect")["Style"]!).Remove("StyleKey"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'style.styleKey'.", exception.Message);
    }

    [Fact]
    public void An_undefined_hatch_pattern_is_named_with_its_value()
    {
        JsonObject root = FullDocument();
        ((JsonObject)Primitive(root, "rect")["Style"]!)["HatchPattern"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid hatch pattern 99.", exception.Message);
    }

    [Fact]
    public void An_undefined_stroke_dash_is_named_with_its_value()
    {
        JsonObject root = FullDocument();
        ((JsonObject)Primitive(root, "rect")["Style"]!)["StrokeDash"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid stroke dash 99.", exception.Message);
    }

    [Fact]
    public void An_unparseable_colour_names_the_style_field()
    {
        JsonObject root = FullDocument();
        ((JsonObject)Primitive(root, "rect")["Style"]!)["FillColour"] = "not-a-colour";

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid colour in style.fillColour.", exception.Message);
    }

    [Fact]
    public void An_undefined_entity_type_is_named_with_its_value()
    {
        JsonObject root = FullDocument();
        Primitive(root, "rect")["EntityType"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid entity type 99.", exception.Message);
    }

    [Fact]
    public void A_text_primitive_without_text_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "text").Remove("Text"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'text'.", exception.Message);
    }

    [Fact]
    public void A_text_primitive_without_bounds_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "text").Remove("TextBounds"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'textBounds'.", exception.Message);
    }

    [Fact]
    public void An_undefined_style_text_alignment_is_named_with_its_value()
    {
        // The style's optional alignment is validated by the reader; the text
        // primitive's own alignment is validated by the SceneText constructor
        // instead, so this fixture must corrupt the style field.
        JsonObject root = FullDocument();
        ((JsonObject)Primitive(root, "text")["Style"]!)["Alignment"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Invalid text alignment 99.", exception.Message);
    }

    [Fact]
    public void An_undefined_primitive_text_alignment_is_refused_as_invalid_scene_data()
    {
        // The primitive's own alignment is cast before SceneText validates it,
        // so the reader reports the wrapped form rather than a field message.
        JsonObject root = FullDocument();
        Primitive(root, "text")["Alignment"] = 99;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Scene JSON contains invalid scene data.", exception.Message);
    }

    [Fact]
    public void A_line_without_a_start_point_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "line").Remove("From"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'from'.", exception.Message);
    }

    [Fact]
    public void A_line_without_an_end_point_names_the_field()
    {
        JsonObject root = FullDocument();
        Assert.True(Primitive(root, "line").Remove("To"));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'to'.", exception.Message);
    }

    [Fact]
    public void A_polygon_with_a_null_point_slot_names_the_field()
    {
        JsonObject root = FullDocument();
        Primitive(root, "polygon")["Points"] = new JsonArray { null };

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Missing scene field 'point'.", exception.Message);
    }

    [Fact]
    public void An_unknown_primitive_kind_is_named_without_echoing_it()
    {
        JsonObject root = FullDocument();
        Primitive(root, "rect")["Kind"] = "nope";

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(Serialized(root)));

        Assert.Equal("Unknown scene primitive kind 'nope'.", exception.Message);
    }

    [Fact]
    public void A_warning_with_an_unknown_owner_kind_is_refused_without_echoing_the_kind()
    {
        // The refusal deliberately does not echo the kind: the value is
        // untrusted input and the message is user-visible.
        JsonNode node = JsonNode.Parse(CanonicalJson())!;
        node["Warnings"] = new JsonArray(
            new JsonObject
            {
                ["Owner"] = new JsonObject { ["Kind"] = "sideways", ["Value"] = "x" },
                ["Code"] = "scene.test",
                ["Message"] = "test warning",
            });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => SceneSnapshot.Deserialize(node.ToJsonString()));

        Assert.Equal("Invalid scene owner.", exception.Message);
    }
}
