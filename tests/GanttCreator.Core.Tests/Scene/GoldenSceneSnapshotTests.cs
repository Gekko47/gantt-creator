using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// Tests for the R3.12 golden scene snapshot.
/// </summary>
/// <remarks>
/// The snapshot is a *scene* baseline derived from the canonical neutral fixture, not
/// a golden image, so it needs no human image review. The change policy is the same
/// as for an image: regenerating it is a dedicated commit with a stated reason, and
/// the comparison is byte-for-byte. Per the work item's stop condition, a failure
/// here means the build is nondeterministic and the cause must be found - the
/// comparison is never loosened.
/// </remarks>
public sealed class GoldenSceneSnapshotTests
{
    /// <summary>Gets the absolute path to the committed golden snapshot.</summary>
    private static string GoldenPath =>
        Path.Combine(AppContext.BaseDirectory, "golden", "scene", "reference-scene.json");

    [Fact]
    public void The_committed_golden_snapshot_exists()
    {
        Assert.True(File.Exists(GoldenPath), "Missing golden scene snapshot: " + GoldenPath);
    }

    [Fact]
    public void The_built_fixture_scene_matches_the_committed_golden_snapshot()
    {
        string built = SceneSnapshot.Serialize(SceneBuilderTests.BuildScene());

        // Regeneration is EXPLICIT, never silent: set GANTTCREATOR_GOLDEN_OUT to the
        // repository's golden file and this fact rewrites it, so a rule change is
        // committed with its reason while an accidental change still fails here. It runs
        // BEFORE the read so a regeneration run compares the file it just wrote.
        RegenerateIfRequested(built);

        string golden = ReadGolden();
        Assert.Equal(golden, built);
    }

    /// <summary>Rewrites the golden file when the regeneration path is supplied.</summary>
    /// <param name="built">The scene serialisation the current rules produce.</param>
    /// <remarks>
    /// The path is an environment variable rather than a constant because the test's own
    /// directory is a build output copy: a constant would rewrite the copy and leave the
    /// repository's golden stale behind a passing run.
    /// </remarks>
    private static void RegenerateIfRequested(string built)
    {
        if (Environment.GetEnvironmentVariable("GANTTCREATOR_GOLDEN_OUT") is { Length: > 0 } path)
        {
            File.WriteAllText(path, built, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }
    }

    [Fact]
    public void Rebuilding_the_scene_twice_produces_byte_identical_output()
    {
        string first = SceneSnapshot.Serialize(SceneBuilderTests.BuildScene());
        string second = SceneSnapshot.Serialize(SceneBuilderTests.BuildScene());

        Assert.Equal(first, second);
    }

    [Fact]
    public void Shuffled_fixture_rows_produce_an_identical_scene()
    {
        // Three orderings, including a full reversal, must all produce the same
        // scene: the snapshot is ordered by the R3.3 z-order keys, never by input
        // order, so a shuffle cannot legitimately change a single byte.
        IReadOnlyList<GanttRowDto> rows = ReferenceSceneFixture.LoadRows();
        string baseline = SceneSnapshot.Serialize(SceneBuilderTests.BuildScene(rows));

        Assert.Equal(baseline, SceneSnapshot.Serialize(SceneBuilderTests.BuildScene([.. rows.Reverse()])));
        Assert.Equal(
            baseline,
            SceneSnapshot.Serialize(SceneBuilderTests.BuildScene([.. rows.OrderBy(row => row.StartCell.Value?.DayNumber ?? 0)])));
    }

    [Fact]
    public void The_golden_snapshot_round_trips_through_deserialization()
    {
        // A snapshot nobody can read back is not a baseline. The round trip also
        // proves the committed file matches the current SceneSnapshot version.
        GanttScene scene = SceneSnapshot.Deserialize(ReadGolden());

        Assert.Equal(ReadGolden(), SceneSnapshot.Serialize(scene));
    }

    [Fact]
    public void The_committed_scene_deserializes_to_the_fixture_built_scene()
    {
        // R8.4 F2's second direction. The byte round trip above proves the DOCUMENT
        // survives deserialize-then-re-serialize, which catches a document field the
        // FromDocument mapping drops or mis-maps. It cannot catch the other direction:
        // a model field that was never added to the snapshot document, because both
        // serializations omit it and the byte comparison still passes. Comparing the
        // deserialized scene to the fixture-built scene object-wise does catch it: a
        // field the document omits deserializes to its default while the built scene
        // carries the fixture's value, so the primitive records differ and this fact
        // fails loudly instead of the field silently defaulting in every raster
        // baseline (the R8.4 fixture route consumes this committed document).
        AssertStructurallyEqual(SceneBuilderTests.BuildScene(), SceneSnapshot.Deserialize(ReadGolden()), "scene");
    }

    private static string ReadGolden() =>
        File.ReadAllText(GoldenPath).TrimEnd('\r', '\n');

    /// <summary>
    /// Asserts two scene values are structurally equal member by member, including
    /// members the snapshot document does not carry.
    /// </summary>
    /// <param name="expected">The fixture-built value.</param>
    /// <param name="actual">The deserialized value.</param>
    /// <param name="path">The member path named in a failure message.</param>
    /// <remarks>
    /// Record equality cannot carry this comparison: a primitive's <c>Points</c> and
    /// a shared owner's <c>OwnedRows</c> are collection members, which record equality
    /// compares by reference. This walk compares scalar leaves through
    /// <see cref="object.Equals(object, object)"/> and recurses into every other
    /// public property and every collection element, so a future member of any shape
    /// is compared by construction rather than through a hand-maintained list.
    /// </remarks>
    private static void AssertStructurallyEqual(object? expected, object? actual, string path)
    {
        if (expected is null || actual is null)
        {
            Assert.True(
                expected is null && actual is null,
                $"{path}: one side is null (expected '{expected ?? "<null>"}', actual '{actual ?? "<null>"}').");
            return;
        }

        Type type = expected.GetType();
        Assert.True(type == actual.GetType(), $"{path}: expected type {type.Name}, actual {actual.GetType().Name}.");

        if (IsScalarLeaf(type))
        {
            Assert.True(Equals(expected, actual), $"{path}: expected '{expected}', actual '{actual}'.");
            return;
        }

        if (expected is System.Collections.IEnumerable expectedItems)
        {
            List<object?> expectedList = [.. expectedItems.Cast<object?>()];
            List<object?> actualList = [.. ((System.Collections.IEnumerable)actual).Cast<object?>()];
            Assert.True(
                expectedList.Count == actualList.Count,
                $"{path}: expected {expectedList.Count} item(s), actual {actualList.Count}.");
            for (int index = 0; index < expectedList.Count; index++)
            {
                AssertStructurallyEqual(expectedList[index], actualList[index], $"{path}[{index}]");
            }

            return;
        }

        foreach (System.Reflection.PropertyInfo property in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            if (property.GetIndexParameters().Length > 0)
            {
                continue;
            }

            AssertStructurallyEqual(property.GetValue(expected), property.GetValue(actual), $"{path}.{property.Name}");
        }
    }

    private static bool IsScalarLeaf(Type type) =>
        type.IsPrimitive
        || type.IsEnum
        || type == typeof(string)
        || type == typeof(decimal)
        || type == typeof(DateOnly)
        || type == typeof(DateTime)
        || type == typeof(TimeSpan)
        || type == typeof(Guid)
        || type == typeof(PointD)
        || type == typeof(RectD)
        || type == typeof(ColourHex)
        || type == typeof(GanttRowId);
}
