using GanttCreator.Core.Scene;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// R4.8A D4: the composition profile decides whether a data panel is drawn, and
/// the live profile is the one that must not draw one.
/// </summary>
/// <remarks>
/// <para>
/// The rule exists because "Live must not draw a data panel" was previously only
/// <em>expressible</em>: <see cref="SceneBuildRequest.Panel"/> is nullable, so a
/// caller could comply by leaving it null — and could equally forget. A live request
/// that carried a panel would draw a replica of the user's own cells on top of them.
/// That failure is silent, and it presents as a layout bug rather than a mistake, so
/// it earns a typed refusal instead of a code-review convention.
/// </para>
/// </remarks>
public class SceneCompositionProfileTests
{
    /// <summary>A panel theme with the five styles the builder requires.</summary>
    private static PanelTheme PanelTheme() =>
        new(
            BodyFill: new SceneStyle("BodyFill", fillColour: ColourHex.Parse("#FFFFFF")),
            BodyText: new SceneStyle("BodyText", textColour: ColourHex.Parse("#000000")),
            HeaderFill: new SceneStyle("HeaderFill", fillColour: ColourHex.Parse("#E6E6E6")),
            HeaderText: new SceneStyle("HeaderText", textColour: ColourHex.Parse("#000000"), fontSizePt: 9, bold: true),
            Border: new SceneStyle("Border", strokeColour: ColourHex.Parse("#BFBFBF"), outlineWidthPt: 0.5));

    /// <summary>The live profile is the only one whose destination already has the table.</summary>
    [Fact]
    public void Only_the_live_profile_does_not_draw_a_data_panel()
    {
        Assert.False(SceneCompositionProfiles.DrawsDataPanel(SceneCompositionProfile.LiveExcel));
        Assert.True(SceneCompositionProfiles.DrawsDataPanel(SceneCompositionProfile.EditableExport));
        Assert.True(SceneCompositionProfiles.DrawsDataPanel(SceneCompositionProfile.PowerPoint));
        Assert.True(SceneCompositionProfiles.DrawsDataPanel(SceneCompositionProfile.Raster));
    }

    /// <summary>
    /// An unrecognised profile value fails <b>closed</b>: it does not draw a panel.
    /// </summary>
    /// <remarks>
    /// The asymmetry is deliberate and load-bearing. Defaulting an unknown profile to
    /// "draws a panel" would hand the destructive behaviour to every future profile
    /// added to the enum without a rule here, and on the live sheet that means a
    /// drawn replica over the user's own cells. A profile nobody has classified is
    /// not yet known to have a panel, so it gets none until someone decides.
    /// </remarks>
    [Fact]
    public void An_unrecognised_profile_value_draws_no_panel()
    {
        Assert.False(SceneCompositionProfiles.DrawsDataPanel((SceneCompositionProfile)9999));
    }

    /// <summary>
    /// The validator's positive test: a live request carrying a panel is refused.
    /// </summary>
    /// <remarks>
    /// This is the case the refusal exists for, constructed deliberately. Without it
    /// the new <c>if</c> would be a branch nothing exercises — the failure mode this
    /// repository has repeatedly shipped.
    /// </remarks>
    [Fact]
    public void A_live_request_that_carries_a_data_panel_is_refused()
    {
        SceneBuildRequest request = new() { Panel = PanelTheme() };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.False(outcome.Succeeded);
        Assert.Equal(SceneBuilderRefusal.LiveProfileCarriesPanel, outcome.Refusal);
    }

    /// <summary>
    /// The same request for a profile that does draw a panel is <b>not</b> refused by
    /// this rule, so the guard is not simply refusing every request with a panel.
    /// </summary>
    /// <remarks>
    /// Without this counterweight, a guard written as "refuse any request carrying a
    /// panel" would pass the test above while breaking every export profile. This is
    /// what makes the rule specific rather than merely restrictive.
    /// </remarks>
    [Fact]
    public void A_non_live_request_may_carry_a_data_panel()
    {
        SceneBuildRequest request =
            new() { Panel = PanelTheme(), Profile = SceneCompositionProfile.PowerPoint };

        SceneBuildOutcome outcome = SceneBuilder.TryBuild(request);

        Assert.NotEqual(SceneBuilderRefusal.LiveProfileCarriesPanel, outcome.Refusal);
    }

    /// <summary>
    /// The profile defaults to live, so a caller who forgets to set one gets the safe
    /// direction rather than a silent panel.
    /// </summary>
    [Fact]
    public void A_request_which_names_no_profile_defaults_to_live()
    {
        Assert.Equal(SceneCompositionProfile.LiveExcel, new SceneBuildRequest().Profile);
    }

    /// <summary>
    /// A live request that omits the panel is unaffected — the common case, and the
    /// one that must not regress when the guard was added.
    /// </summary>
    /// <remarks>
    /// The request is otherwise empty, so it is refused for a missing measurement.
    /// The assertion is that the refusal is <em>not</em> the panel rule, which is
    /// what "unaffected" can mean for an incomplete request.
    /// </remarks>
    [Fact]
    public void A_live_request_without_a_panel_is_not_refused_by_the_panel_rule()
    {
        SceneBuildOutcome outcome = SceneBuilder.TryBuild(new SceneBuildRequest());

        Assert.NotEqual(SceneBuilderRefusal.LiveProfileCarriesPanel, outcome.Refusal);
    }
}
