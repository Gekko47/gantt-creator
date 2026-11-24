using GanttCreator.Core.Scene;
using GanttCreator.Raster;

namespace GanttCreator.AddIn;

/// <summary>The outcome of checking a scene's fonts against the raster engine (R8.8 D10).</summary>
/// <param name="Scene">The scene to render: the input scene, or a substituted copy.</param>
/// <param name="SubstitutedFamilies">
/// The distinct families that were replaced, ordinal-sorted; empty when none were.
/// </param>
internal sealed record PngFontGateResult(GanttScene Scene, IReadOnlyList<string> SubstitutedFamilies)
{
    /// <summary>Gets whether any family was substituted, so the caller must report it.</summary>
    internal bool Substituted => SubstitutedFamilies.Count > 0;
}

/// <summary>
/// Identifies whether the scene's text families resolve through the raster engine's pinned
/// resolver and builds a substituted scene when they do not (R8.8 D10, ADR-0070 D3 amendment).
/// </summary>
/// <remarks>
/// <para>
/// <b>Identify, then substitute with a report, never fail on the user.</b> The pinned faces are
/// the only faces the engine resolves, so the check is one
/// <see cref="PinnedFontSource.Resolve"/> call per distinct family and weight the scene names.
/// When a family cannot be resolved but the pinned faces themselves load, the scene's text
/// styles are rebuilt with the substitute family through the validated
/// <see cref="SceneStyle"/> constructor and the substitution is named in the result; when the
/// pinned set itself is unavailable there is nothing to substitute with, the ORIGINAL scene is
/// returned unchanged, and the renderer's typed refusal stands.
/// </para>
/// <para>
/// <b>Pure and host-free.</b> The resolver is injected, so a test drives every arm without
/// Office fonts.
/// </para>
/// </remarks>
internal static class PngFontGate
{
    /// <summary>Resolves a family/weight pair (injectable for tests).</summary>
    /// <param name="fontFamily">The scene style's family name.</param>
    /// <param name="bold">Whether the style is bold.</param>
    /// <returns>The resolution outcome; a refusal is a value, never a throw.</returns>
    internal delegate PinnedFontResolution FamilyResolver(string fontFamily, bool bold);

    /// <summary>Checks the scene's families and substitutes the pinned family when needed.</summary>
    /// <param name="scene">The export scene.</param>
    /// <param name="resolver">The pinned-font resolver.</param>
    /// <param name="substituteFamily">The pinned label family to substitute.</param>
    /// <returns>The gate result; the original scene when nothing was substituted.</returns>
    internal static PngFontGateResult Apply(
        GanttScene scene,
        FamilyResolver resolver,
        string substituteFamily)
    {
        ArgumentNullException.ThrowIfNull(scene);
        ArgumentNullException.ThrowIfNull(resolver);
        ArgumentException.ThrowIfNullOrWhiteSpace(substituteFamily);

        SortedSet<string> failed = new(StringComparer.Ordinal);
        foreach (SceneText text in scene.Primitives.OfType<SceneText>())
        {
            // SceneText stores its style through WithLabelFontFamilyIfMissing, so the family
            // the scene names is never blank here; a resolver refusal is a value, never a throw.
            string family = text.Style.FontFamily!;
            if (!resolver(family, text.Style.Bold == true).Succeeded)
            {
                _ = failed.Add(family);
            }
        }

        if (failed.Count == 0)
        {
            return new PngFontGateResult(scene, []);
        }

        GanttScene substituted = Substitute(scene, substituteFamily);

        // The substitution is only honest when the pinned faces themselves load for every
        // weight it uses. If one does not, there is nothing to render with: the original
        // scene goes back and the renderer's typed refusal names the missing pinned face.
        foreach (SceneText text in substituted.Primitives.OfType<SceneText>())
        {
            if (!resolver(text.Style.FontFamily!, text.Style.Bold == true).Succeeded)
            {
                return new PngFontGateResult(scene, []);
            }
        }

        return new PngFontGateResult(substituted, [.. failed]);
    }

    /// <summary>Rebuilds the scene's text primitives with the substitute family.</summary>
    /// <param name="scene">The scene whose text styles are rewritten.</param>
    /// <param name="substituteFamily">The family every text style names afterwards.</param>
    /// <returns>A new scene; every non-text primitive is carried over unchanged.</returns>
    private static GanttScene Substitute(GanttScene scene, string substituteFamily)
    {
        List<ScenePrimitive> primitives = [];
        foreach (ScenePrimitive primitive in scene.Primitives)
        {
            if (primitive is not SceneText text)
            {
                primitives.Add(primitive);
                continue;
            }

            SceneStyle style = new(
                text.Style.StyleKey,
                text.Style.FillColour,
                text.Style.StrokeColour,
                text.Style.OutlineWidthPt,
                text.Style.HatchPattern,
                substituteFamily,
                text.Style.FontSizePt,
                text.Style.Bold,
                text.Style.Alignment,
                text.Style.TextColour,
                text.Style.StrokeDash);
            primitives.Add(new SceneText(
                text.PrimitiveId,
                text.OwnerId,
                text.ZLayer,
                text.Text,
                text.TextBounds,
                style,
                text.Alignment,
                text.EntityType,
                text.LaneOrder,
                text.StackIndex,
                text.SortOrder,
                text.LeftInsetPt));
        }

        return GanttScene.TryCreate(scene.ChartBounds, scene.PlotBounds, primitives, scene.Warnings).Scene!;
    }
}
