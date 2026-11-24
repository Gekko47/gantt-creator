using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Raster;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// R8.8 D10, ADR-0070 D3 amendment: the export identifies font compatibility with the raster
/// engine and substitutes the pinned label family with a report. The three arms are the
/// contract — everything resolves (the scene is untouched), a family the engine cannot resolve
/// (only the text styles are rewritten, the substitution is named), and the pinned set itself
/// missing (nothing to substitute with, so the original scene and the renderer's typed
/// refusal stand). The resolver is injected, so every arm runs host-free.
/// </summary>
public sealed class PngFontGateTests
{
    private const string Substitute = "Aptos";

    private static readonly RectD Bounds = new(0, 0, 400, 200);
    private static readonly string[] TwoFamilies = ["Comic Sans MS", "Wingdings"];
    private static readonly string[] OneFamily = ["Comic Sans MS"];

    private static GanttScene Scene(params ScenePrimitive[] primitives) =>
        GanttScene.TryCreate(Bounds, Bounds, primitives, []).Scene!;

    private static SceneText Text(string id, string family, bool bold = false) =>
        new(
            id,
            SceneOwnerId.Chart,
            ZLayer.Title,
            "T",
            new RectD(10, 10, 50, 10),
            new SceneStyle("S-" + id, fontFamily: family, fontSizePt: 9, bold: bold),
            GanttTextAlignment.Centre);

    private static SceneRect Rect(string id) =>
        new(
            id,
            SceneOwnerId.Chart,
            ZLayer.ActivityBody,
            new RectD(20, 20, 40, 12),
            new SceneStyle("R-" + id, fillColour: ColourHex.Parse("#010203")));

    /// <summary>A resolution that reports the face as renderable (a host-free stand-in).</summary>
    private static PinnedFontResolution Resolved(string family) =>
        PinnedFontResolution.Resolved(
            null!,
            new PinnedFontObservation(family, family, string.Empty, double.NaN, 1000, null, string.Empty));

    /// <summary>A resolution that reports the face as unavailable.</summary>
    private static PinnedFontResolution Refused() =>
        PinnedFontResolution.Refused(PinnedFontRefusal.FontFileMissing, null);

    /// <summary>Arm 1: every family resolves, so the input scene is returned unchanged.</summary>
    [Fact]
    public void Apply_returns_the_input_scene_when_every_family_resolves()
    {
        GanttScene scene = Scene(Text("a", "Aptos"), Text("b", "Aptos Narrow", bold: true));

        PngFontGateResult result = PngFontGate.Apply(scene, (family, _) => Resolved(family), Substitute);

        Assert.Same(scene, result.Scene);
        Assert.False(result.Substituted);
        Assert.Empty(result.SubstitutedFamilies);
    }

    /// <summary>
    /// Arm 2: an unresolvable family is substituted for every text style, the substitution is
    /// reported (ordinal-sorted), every other style fact is preserved, and the non-text
    /// primitives are carried over unchanged.
    /// </summary>
    [Fact]
    public void Apply_substitutes_an_unresolvable_family_and_reports_it()
    {
        GanttScene scene = Scene(Text("a", "Comic Sans MS"), Rect("r"), Text("b", "Wingdings", bold: true));

        PngFontGateResult result = PngFontGate.Apply(
            scene,
            (family, _) => string.Equals(family, Substitute, StringComparison.Ordinal)
                ? Resolved(family)
                : Refused(),
            Substitute);

        Assert.True(result.Substituted);
        Assert.NotSame(scene, result.Scene);
        Assert.Equal(TwoFamilies, result.SubstitutedFamilies);
        Assert.All(
            result.Scene.Primitives.OfType<SceneText>(),
            text => Assert.Equal(Substitute, text.Style.FontFamily));

        SceneText regular = result.Scene.Primitives.OfType<SceneText>().Single(text => text.PrimitiveId == "a");
        Assert.Equal(9, regular.Style.FontSizePt);
        Assert.Equal(GanttTextAlignment.Centre, regular.Alignment);

        SceneText bold = result.Scene.Primitives.OfType<SceneText>().Single(text => text.PrimitiveId == "b");
        Assert.True(bold.Style.Bold);

        SceneRect preserved = result.Scene.Primitives.OfType<SceneRect>().Single();
        Assert.Equal("r", preserved.PrimitiveId);
    }

    /// <summary>
    /// Arm 3: the pinned faces themselves do not load, so there is nothing to substitute with;
    /// the original scene goes back untouched and the renderer's typed refusal stands.
    /// </summary>
    [Fact]
    public void Apply_returns_the_input_scene_when_the_pinned_faces_are_missing()
    {
        GanttScene scene = Scene(Text("a", "Comic Sans MS"));

        PngFontGateResult result = PngFontGate.Apply(scene, (_, _) => Refused(), Substitute);

        Assert.Same(scene, result.Scene);
        Assert.False(result.Substituted);
        Assert.Empty(result.SubstitutedFamilies);
    }

    /// <summary>
    /// The gate is weight-sensitive: a family that resolves at regular but refuses at bold is
    /// what drives the substitution, and the report names that family.
    /// </summary>
    [Fact]
    public void Apply_treats_an_unresolvable_weight_as_a_substitution()
    {
        GanttScene scene = Scene(Text("a", "Comic Sans MS", bold: true));

        PngFontGateResult result = PngFontGate.Apply(
            scene,
            (family, bold) => bold && string.Equals(family, "Comic Sans MS", StringComparison.Ordinal)
                ? Refused()
                : Resolved(family),
            Substitute);

        Assert.True(result.Substituted);
        Assert.Equal(OneFamily, result.SubstitutedFamilies);
    }
}
