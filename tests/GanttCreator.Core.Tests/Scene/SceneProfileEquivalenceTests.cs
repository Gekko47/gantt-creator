using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R8.4 profile-equivalence pin: <see cref="SceneBuilder"/> emits identical
/// output for the two non-live profiles the raster baselines and the production
/// export rely on.
/// </summary>
/// <remarks>
/// <para>
/// The R8.4 golden baselines render a
/// <see cref="SceneCompositionProfile.Raster"/>-profile fixture scene, while R8.8's
/// production scene is <see cref="SceneCompositionProfile.EditableExport"/> through
/// <c>ExportBoundsResolver</c>. Both are non-live compositions, and the builder's
/// profile gates test only <c>== LiveExcel</c> (the drawn title, the live panel
/// bounds, and sheet coordinates), so the two profiles are identical by
/// construction. This pin makes the equivalence proven rather than assumed;
/// without it the baselines and the production scene would differ by an unproven
/// profile label (the Phase-8 review's F3, promoted to a prerequisite of R8.4).
/// </para>
/// <para>
/// The third fact is the sensitivity control. A pin that compares two scenes can
/// pass vacuously if the comparison cannot see a profile-driven difference at all,
/// so it proves the <see cref="SceneCompositionProfile.LiveExcel"/> profile really
/// does change the output for the same request: the drawn title exists only
/// outside the live profile (ADR-0030 D6).
/// </para>
/// </remarks>
public sealed class SceneProfileEquivalenceTests
{
    [Fact]
    public void The_raster_and_editable_export_profiles_compose_an_identical_chart_scene()
    {
        SceneBuildRequest raster = SceneBuilderTests.BuildRequest();

        Assert.Equal(
            SceneSnapshot.Serialize(Build(raster)),
            SceneSnapshot.Serialize(Build(raster with { Profile = SceneCompositionProfile.EditableExport })));
    }

    [Fact]
    public void The_raster_and_editable_export_profiles_compose_an_identical_panel_scene()
    {
        SceneBuildRequest raster = SceneBuilderTests.BuildRequest(withPanel: true);

        Assert.Equal(
            SceneSnapshot.Serialize(Build(raster)),
            SceneSnapshot.Serialize(Build(raster with { Profile = SceneCompositionProfile.EditableExport })));
    }

    [Fact]
    public void The_live_profile_composes_a_different_scene_so_the_pin_can_detect_a_profile_change()
    {
        SceneBuildRequest raster = SceneBuilderTests.BuildRequest();

        Assert.NotEqual(
            SceneSnapshot.Serialize(Build(raster)),
            SceneSnapshot.Serialize(Build(raster with { Profile = SceneCompositionProfile.LiveExcel })));
    }

    private static GanttScene Build(SceneBuildRequest request)
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);
        Assert.True(outcome.Succeeded, "Profile-equivalence build refused: " + outcome.Refusal);
        return outcome.Result!.Scene;
    }
}
