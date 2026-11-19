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

    private static string ReadGolden() =>
        File.ReadAllText(GoldenPath).TrimEnd('\r', '\n');
}
