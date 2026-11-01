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
