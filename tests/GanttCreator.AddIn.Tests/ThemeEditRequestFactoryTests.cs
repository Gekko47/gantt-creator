using GanttCreator.Core;
using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The theme dialog's opening request (audit finding 4). It must show what the chart currently
/// draws: a stored colour wins over the preset, a preset with no stored formatting keeps its own
/// colours, and only the one user-editable line type is projected from the workbook.
/// </summary>
/// <remarks>
/// This projection used to be the larger half of <c>ShowThemeSettingsDialog</c>, a 187-line method
/// on the state service, and had no test of its own.
/// </remarks>
public class ThemeEditRequestFactoryTests
{
    private static readonly GanttLabelPosition _anyPosition =
        Enum.GetValues<GanttLabelPosition>().Last();

    private const GanttStrokeDash _anyDash = GanttStrokeDash.Dash;

    [Fact]
    public void A_stored_colour_wins_over_the_preset()
    {
        GanttStylePreset preset = GanttCatalogues.StylePresets[0];

        ThemeEditRequest request = ThemeEditRequestFactory.FromWorkbook(Outcome(
            Definition(preset.StyleKey, "#123456", "#234567", "#345678", dash: null)));

        ThemeStyleEdit edit = request.StyleEdits.Single(style => style.StyleKey == preset.StyleKey);
        Assert.Equal("#123456", edit.FillColour);
        Assert.Equal("#234567", edit.StrokeColour);
        Assert.Equal("#345678", edit.TextColour);
    }

    [Fact]
    public void A_preset_with_no_stored_formatting_keeps_its_own_colours()
    {
        GanttStylePreset preset = GanttCatalogues.StylePresets[0];

        ThemeEditRequest request = ThemeEditRequestFactory.FromWorkbook(Outcome());

        ThemeStyleEdit edit = request.StyleEdits.Single(style => style.StyleKey == preset.StyleKey);
        Assert.Equal(preset.FillColour, edit.FillColour);
        Assert.Equal(preset.StrokeColour, edit.StrokeColour);
        Assert.Equal(preset.TextColour, edit.TextColour);
    }

    [Fact]
    public void Only_an_editable_styles_line_type_is_projected()
    {
        GanttStylePreset editable = GanttCatalogues.StylePresets
            .First(preset => GanttStyleDashPolicy.IsUserEditable(preset.StyleKey));
        GanttStylePreset fixedDash = GanttCatalogues.StylePresets
            .First(preset => !GanttStyleDashPolicy.IsUserEditable(preset.StyleKey));

        ThemeEditRequest request = ThemeEditRequestFactory.FromWorkbook(Outcome(
            Definition(editable.StyleKey, "#111111", "#222222", "#333333", _anyDash),
            Definition(fixedDash.StyleKey, "#111111", "#222222", "#333333", _anyDash)));

        Assert.Equal(
            _anyDash,
            request.StyleEdits.Single(style => style.StyleKey == editable.StyleKey).StrokeDash);
        Assert.Null(request.StyleEdits.Single(style => style.StyleKey == fixedDash.StyleKey).StrokeDash);
    }

    [Fact]
    public void A_custom_style_is_carried_with_its_display_name()
    {
        ThemeEditRequest request = ThemeEditRequestFactory.FromWorkbook(Outcome(
            Definition("MyCustomStyle", "#111111", "#222222", "#333333", dash: null, displayName: "Field work")));

        CustomStyleDef custom = Assert.Single(request.CustomStyleDefs!);
        Assert.Equal("MyCustomStyle", custom.StyleKey);
        Assert.Equal("Field work", custom.DisplayName);
    }

    [Fact]
    public void A_null_outcome_is_rejected() =>
        _ = Assert.Throws<ArgumentNullException>(() => _ = ThemeEditRequestFactory.FromWorkbook(null!));

    private static GanttStyleDefinition Definition(
        string styleKey,
        string fill,
        string stroke,
        string text,
        GanttStrokeDash? dash,
        string? displayName = null) =>
        new(
            styleKey,
            new HashSet<GanttLabelPosition> { _anyPosition },
            EntityColourCapability.Fill | EntityColourCapability.Stroke,
            DefaultLabelPosition: _anyPosition,
            FillColour: fill,
            StrokeColour: stroke,
            TextColour: text,
            HatchPattern: GanttHatchPattern.None,
            DisplayName: displayName,
            StrokeDash: dash);

    private static ConfigReadOutcome Outcome(params GanttStyleDefinition[] styles) =>
        new(
            Succeeded: true,
            WorkbookId: "wb-test",
            Settings: new Dictionary<string, string>(StringComparer.Ordinal),
            Styles: new GanttStyleRegistry(styles),
            Metrics: new Dictionary<string, double>(StringComparer.Ordinal),
            Refusal: null);
}
