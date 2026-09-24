using System.Globalization;

namespace GanttCreator.Core;

/// <summary>
/// The single code-owned first-release configuration catalogue: the
/// materialised tables of <c>_GanttCreatorConfig</c> and the catalogue hash.
/// Every value transcribes the entity guide (revision 3, 2026-09-21) and is
/// immutable; the Office writer projects this catalogue deterministically
/// (ADR-0007). No other component maintains a second list.
/// </summary>
/// <remarks>
/// <para>
/// Table names, anchors, and header rows are part of the workbook schema:
/// renaming any of them is a schema migration that bumps
/// <see cref="GanttSchemaVersion.CurrentSchemaVersion"/>. The catalogue hash
/// covers the definitions below (never raw worksheet text) so R2.9/R2.10
/// detect drift (ADR-0007 D6).
/// </para>
/// </remarks>
public static class GanttCatalogues
{
    /// <summary>The Excel Table name for the materialised type catalogue.</summary>
    public const string TypesTableName = "tblGanttTypes";

    /// <summary>The Excel Table name for the style presets.</summary>
    public const string StylesTableName = "tblGanttStyles";

    /// <summary>The Excel Table name for the metric tokens.</summary>
    public const string MetricsTableName = "tblGanttMetrics";

    /// <summary>The Excel Table name for the first-release settings.</summary>
    public const string SettingsTableName = "tblGanttSettings";

    /// <summary>The Excel Table name for the add-in metadata (schema version, hash, workbook ID).</summary>
    public const string ConfigTableName = "tblGanttConfig";

    /// <summary>The anchor cell of <see cref="TypesTableName"/> (ADR-0007 D7).</summary>
    public const string TypesAnchor = "A1";

    /// <summary>The anchor cell of <see cref="StylesTableName"/> (ADR-0007 D7).</summary>
    public const string StylesAnchor = "A20";

    /// <summary>The anchor cell of <see cref="MetricsTableName"/> (ADR-0007 D7).</summary>
    public const string MetricsAnchor = "A40";

    /// <summary>The anchor cell of <see cref="SettingsTableName"/> (ADR-0007 D7).</summary>
    public const string SettingsAnchor = "A70";

    /// <summary>The anchor cell of <see cref="ConfigTableName"/> (ADR-0007 D7).</summary>
    public const string ConfigAnchor = "A90";

    /// <summary>The <see cref="ConfigTableName"/> row key for the schema version.</summary>
    public const string ConfigSchemaVersionKey = "SchemaVersion";

    /// <summary>The <see cref="ConfigTableName"/> row key for the catalogue hash.</summary>
    public const string ConfigCatalogueHashKey = "CatalogueHash";

    /// <summary>The <see cref="ConfigTableName"/> row key for the workbook's stable ID.</summary>
    public const string ConfigWorkbookIdKey = "WorkbookId";

    /// <summary>The <see cref="ConfigTableName"/> row key for the add-in version last used.</summary>
    public const string ConfigAddInVersionKey = "AddInVersionLastUsed";

    /// <summary>The <see cref="TypesTableName"/> header row, in contract order.</summary>
    public static readonly string[] TypesHeaders =
    [
        "TypeName",
        "DisplayName",
        "Kind",
        "DateMode",
        "DefaultStyleKey",
        "ColourCapability",
        "RequiresStyleKey",
        "AllowedLabelPositions",
    ];

    /// <summary>The <see cref="StylesTableName"/> header row, in contract order.</summary>
    public static readonly string[] StylesHeaders =
    [
        "StyleKey",
        "DisplayName",
        "FillColour",
        "StrokeColour",
        "HatchPattern",
        "HatchPitchPt",
        "HatchLinePt",
        "TextColour",
        "StandardOutlinePt",
        "ActivityHeightPt",
        "MilestoneSizePt",
        "DefaultLabelPosition",
        "AllowedLabelPositions",
        "ColourCapability",
    ];

    /// <summary>The <see cref="MetricsTableName"/> header row, in contract order.</summary>
    public static readonly string[] MetricsHeaders =
    [
        "MetricName",
        "DefaultValue",
        "Minimum",
        "Maximum",
    ];

    /// <summary>The <see cref="SettingsTableName"/> header row, in contract order.</summary>
    public static readonly string[] SettingsHeaders =
    [
        "SettingKey",
        "SettingValue",
    ];

    /// <summary>The <see cref="ConfigTableName"/> header row, in contract order.</summary>
    public static readonly string[] ConfigHeaders =
    [
        "ConfigKey",
        "ConfigValue",
    ];

    /// <summary>
    /// The 23 metric tokens, in entity-guide "Shared metric tokens" order
    /// (name, default, valid range; all values in points).
    /// </summary>
    public static IReadOnlyList<GanttMetricToken> Metrics { get; } =
    [
        new("ChartOuterPaddingPt", 6, 0, 36),
        new("TitleBandHeightPt", 24, 12, 72),
        new("YearBandHeightPt", 18, 10, 48),
        new("PeriodBandHeightPt", 16, 10, 48),
        new("MinimumHeaderLabelWidthPt", 18, 6, 72),
        new("LaneHeightPt", 18, 10, 72),
        new("SplitterHeightPt", 18, 10, 72),
        new("SpacerHeightPt", 9, 0, 72),
        new("LanePaddingTopPt", 3, 0, 18),
        new("LanePaddingBottomPt", 3, 0, 18),
        new("ActivityHeightPt", 8, 2, 36),
        new("StackGapPt", 2, 0, 12),
        new("MilestoneSizePt", 8, 3, 36),
        new("LabelGapPt", 3, 0, 18),
        new("LabelHeightPt", 10, 6, 36),
        new("MaximumExternalLabelWidthPt", 144, 36, 360),
        new("StandardOutlinePt", 0.75, 0, 6),
        new("CriticalLinePt", 2.25, 0.5, 12),
        new("GridLinePt", 0.5, 0.25, 3),
        new("MajorBoundaryPt", 1, 0.25, 6),
        new("DelineatorLinePt", 0.75, 0.25, 6),
        new("HatchPitchPt", 4, 2, 18),
        new("HatchLinePt", 0.5, 0.25, 3),
    ];

    /// <summary>
    /// The 21 colour tokens, in entity-guide "Shared colour and typography
    /// tokens" order (uppercase <c>#RRGGBB</c>; explicit alpha <c>FF</c>
    /// unless transparency is named).
    /// </summary>
    public static IReadOnlyList<GanttColourToken> Colours { get; } =
    [
        new("ActualFill", "#00B0F0"),
        new("ActualOutline", "#0070C0"),
        new("PlannedFill", "#92D050"),
        new("PlannedOutline", "#548235"),
        new("BaselineFill", "#00B050"),
        new("BaselineOutline", "#006100"),
        new("CriticalStroke", "#FF0000"),
        new("CriticalOutline", "#C00000"),
        new("DelayFill", "#FF0000"),
        new("DelayText", "#FFFFFF"),
        new("DefaultText", "#000000"),
        new("ChartBackground", "#FFFFFF"),
        new("DataPanelFill", "#FFFFFF"),
        new("HeaderFill", "#FFFFFF"),
        new("YearHeaderFill", "#D9D9D9"),
        new("SplitterFill", "#FFE699"),
        new("AlternateBandFill", "#F2F2F2"),
        new("MinorGridStroke", "#D9D9D9"),
        new("MajorGridStroke", "#000000"),
        new("DelineatorStroke", "#404040"),
        new("WarningFill", "#FFF2CC"),
    ];

    /// <summary>
    /// The 6 typography tokens, in entity-guide order (bold flags for
    /// <c>YearFontSizePt</c>, <c>TitleFontSizePt</c>, and
    /// <c>DelineatorFontSizePt</c>).
    /// </summary>
    public static IReadOnlyList<GanttTypographyToken> Typography { get; } =
    [
        new("FontFamily", "Aptos", isBold: false),
        new("BodyFontSizePt", "8", isBold: false),
        new("HeaderFontSizePt", "8", isBold: false),
        new("YearFontSizePt", "9", isBold: true),
        new("TitleFontSizePt", "11", isBold: true),
        new("DelineatorFontSizePt", "8", isBold: true),
    ];

    /// <summary>
    /// The 11 first-release setting definitions approved in ADR-0007 D3.
    /// <c>PlotStartMode</c>/<c>PlotFinishMode</c>/<c>TimeScale</c> value
    /// enums are fixed by R3.3/R5.1; until then the defaults below are the
    /// only stored values.
    /// </summary>
    public static IReadOnlyList<GanttSettingDefinition> Settings { get; } =
    [
        new("ChartTitle", string.Empty),
        new("ShowTitle", "TRUE"),
        new("TimeScale", "Month"),
        new("LegendPosition", "Right"),
        new("ExportIncludeDataPanel", "TRUE"),
        new("ExportIncludeLegend", "TRUE"),
        new("PlotStartMode", "DataRange"),
        new("PlotFinishMode", "DataRange"),
        new("AlternateBanding", "TRUE"),
        new("ShowMinorGrid", "TRUE"),
        new("ShowMajorGrid", "TRUE"),
    ];

    private static readonly Dictionary<string, Func<GanttStylePreset>> _resolvers = new(StringComparer.Ordinal)
    {
        ["Splitter"] = () => Preset(
            "Splitter", "Splitter", "SplitterFill", null, GanttHatchPattern.None, "DefaultText",
            null, "SplitterHeightPt", null),
        ["Spacer"] = () => Preset(
            "Spacer", "Spacer", null, null, GanttHatchPattern.None, "DefaultText",
            null, "SpacerHeightPt", null),
        ["AsBuiltActivity"] = () => Preset(
            "AsBuiltActivity", "As-Built Activity", "ActualFill", "ActualOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["AsPlannedActivity"] = () => Preset(
            "AsPlannedActivity", "As-Planned Activity", "PlannedFill", "PlannedOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["BaselineActivity"] = () => Preset(
            "BaselineActivity", "Baseline Activity", "BaselineFill", "BaselineOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["CriticalInterval"] = () => Preset(
            "CriticalInterval", "Critical Interval", null, "CriticalStroke",
            GanttHatchPattern.None, "DefaultText", "CriticalLinePt", "ActivityHeightPt", null),
        ["DelayEvent"] = () => Preset(
            "DelayEvent", "Delay Event", "DelayFill", "CriticalOutline",
            GanttHatchPattern.None, "DelayText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["AsBuiltProcurement"] = () => Preset(
            "AsBuiltProcurement", "As-Built Procurement", "ActualFill", "ActualOutline",
            GanttHatchPattern.ForwardDiagonal, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["AsPlannedProcurement"] = () => Preset(
            "AsPlannedProcurement", "As-Planned Procurement", "PlannedFill", "PlannedOutline",
            GanttHatchPattern.ForwardDiagonal, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["BaselineProcurement"] = () => Preset(
            "BaselineProcurement", "Baseline Procurement", "BaselineFill", "BaselineOutline",
            GanttHatchPattern.ForwardDiagonal, "DefaultText", "StandardOutlinePt", "ActivityHeightPt", null),
        ["AsBuiltMilestone"] = () => Preset(
            "AsBuiltMilestone", "As-Built Milestone", "ActualFill", "ActualOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", null, "MilestoneSizePt"),
        ["AsPlannedMilestone"] = () => Preset(
            "AsPlannedMilestone", "As-Planned Milestone", "PlannedFill", "PlannedOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", null, "MilestoneSizePt"),
        ["BaselineMilestone"] = () => Preset(
            "BaselineMilestone", "Baseline Milestone", "BaselineFill", "BaselineOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", null, "MilestoneSizePt"),
        ["CriticalMilestone"] = () => Preset(
            "CriticalMilestone", "Critical Milestone", "CriticalStroke", "CriticalOutline",
            GanttHatchPattern.None, "DefaultText", "StandardOutlinePt", null, "MilestoneSizePt"),
        ["DefaultDelineator"] = () => Preset(
            "DefaultDelineator", "Delineator", null, "DelineatorStroke",
            GanttHatchPattern.None, "DefaultText", "DelineatorLinePt", null, null),
    };

    /// <summary>
    /// The built-in style presets, in <see cref="EntityTypeCatalog.Entries"/>
    /// order: one row per entry whose <c>DefaultStyleKey</c> is non-empty
    /// (15 of 16; <c>Custom Activity</c> requires a user <c>StyleKey</c>).
    /// Each preset resolves its colours and metrics from the token defaults
    /// above (ADR-0007 D7).
    /// </summary>
    public static IReadOnlyList<GanttStylePreset> StylePresets { get; } =
        [
            .. EntityTypeCatalog.Entries
                .Where(entry => entry.DefaultStyleKey.Length > 0)
                .Select(entry => GetPreset(entry.DefaultStyleKey)),
        ];

    /// <summary>
    /// The projected <c>tblGanttTypes</c> rows, in
    /// <see cref="EntityTypeCatalog.Entries"/> order (16 rows). Allowed
    /// label positions are ordinal-sorted for deterministic serialisation;
    /// an empty set serialises as empty (the position is resolved from the
    /// row's required named style).
    /// </summary>
    public static IReadOnlyList<GanttTypeCatalogueRow> TypeRows { get; } =
        [
            .. EntityTypeCatalog.Entries
                .Select(entry => new GanttTypeCatalogueRow(
                    TypeName: entry.Type.ToString(),
                    DisplayName: entry.DisplayName,
                    Kind: entry.Kind.ToString(),
                    DateMode: entry.DateMode.ToString(),
                    DefaultStyleKey: entry.DefaultStyleKey,
                    ColourCapability: entry.ColourCapability.ToString(),
                    RequiresStyleKey: entry.RequiresStyleKey ? "TRUE" : "FALSE",
                    AllowedLabelPositions: string.Join(
                        " ",
                        entry.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)))),
        ];

    /// <summary>
    /// Resolves the built-in style preset for a style key from the token
    /// defaults. The mapping transcribes the entity guide's per-type default
    /// styles; an unknown key is not a preset and throws (the guard's
    /// positive test lives in <c>GanttConfigCatalogueTests</c>).
    /// </summary>
    /// <param name="styleKey">The exact style key.</param>
    /// <returns>The resolved preset.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="styleKey"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Thrown when the style key is not a built-in preset key.
    /// </exception>
    public static GanttStylePreset GetPreset(string styleKey) =>
        _resolvers.TryGetValue(styleKey, out Func<GanttStylePreset>? resolve)
            ? resolve()
            : throw new ArgumentOutOfRangeException(
                nameof(styleKey),
                styleKey,
                "Style key is not a built-in preset key.");

    /// <summary>
    /// Resolves one preset from token names; a null token name means the
    /// column is not applicable and resolves to empty/0 (documented in
    /// ADR-0007 D7).
    /// </summary>
    private static GanttStylePreset Preset(
        string styleKey,
        string displayName,
        string? fillToken,
        string? strokeToken,
        GanttHatchPattern hatchPattern,
        string textToken,
        string? outlineToken,
        string? heightToken,
        string? milestoneToken)
    {
        EntityTypeDefinition definition = EntityTypeCatalog.Entries.First(entry =>
            string.Equals(entry.DefaultStyleKey, styleKey, StringComparison.Ordinal));
        GanttLabelPosition defaultLabelPosition = styleKey switch
        {
            "DelayEvent" => GanttLabelPosition.Inside,
            "CriticalInterval" or "Spacer" => GanttLabelPosition.None,
            "Splitter" => GanttLabelPosition.DataPanelLeft,
            _ => GanttLabelPosition.Auto,
        };
        var outline = outlineToken is null ? 0 : ResolveMetric(outlineToken);
        var height = heightToken is null ? 0 : ResolveMetric(heightToken);
        var milestone = milestoneToken is null ? 0 : ResolveMetric(milestoneToken);
        return new GanttStylePreset(
            styleKey,
            displayName,
            fillToken is null ? string.Empty : ResolveColour(fillToken),
            strokeToken is null ? string.Empty : ResolveColour(strokeToken),
            hatchPattern,
            ResolveMetric("HatchPitchPt"),
            ResolveMetric("HatchLinePt"),
            ResolveColour(textToken),
            outline,
            height,
            milestone,
            defaultLabelPosition,
            definition.AllowedLabelPositions,
            definition.ColourCapability);
    }

    private static string ResolveColour(string tokenName) =>
        Colours.First(token => token.Name == tokenName).HexValue;

    private static double ResolveMetric(string tokenName) =>
        Metrics.First(token => token.Name == tokenName).DefaultValue;

    /// <summary>
    /// Computes the catalogue hash (ADR-0007 D6): SHA-256 over the canonical,
    /// culture-invariant serialisation of the code-owned catalogue
    /// definitions, as lowercase hex. Never computed over worksheet text.
    /// </summary>
    /// <returns>The 64-character lowercase hexadecimal digest.</returns>
    public static string ComputeCatalogueHash()
    {
        var builder = new System.Text.StringBuilder();
        AppendLine(builder, $"schema|{GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture)}");
        AppendTableHeader(builder, TypesTableName, TypesHeaders);
        foreach (GanttTypeCatalogueRow row in TypeRows)
        {
            AppendLine(
                builder,
                Join(row.TypeName, row.DisplayName, row.Kind, row.DateMode, row.DefaultStyleKey,
                    row.ColourCapability, row.RequiresStyleKey, row.AllowedLabelPositions));
        }

        AppendTableHeader(builder, StylesTableName, StylesHeaders);
        foreach (GanttStylePreset preset in StylePresets)
        {
            AppendLine(
                builder,
                Join(
                    preset.StyleKey, preset.DisplayName, preset.FillColour, preset.StrokeColour,
                    preset.HatchPattern.ToString(),
                    Format(preset.HatchPitchPt), Format(preset.HatchLinePt), preset.TextColour,
                    Format(preset.StandardOutlinePt), Format(preset.ActivityHeightPt),
                    Format(preset.MilestoneSizePt),
                    preset.DefaultLabelPosition.ToString(),
                    string.Join(
                        " ",
                        preset.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)),
                    preset.ColourCapability.ToString()));
        }

        AppendTableHeader(builder, MetricsTableName, MetricsHeaders);
        foreach (GanttMetricToken metric in Metrics)
        {
            AppendLine(
                builder,
                Join(metric.Name, Format(metric.DefaultValue), Format(metric.Minimum), Format(metric.Maximum)));
        }

        AppendLine(builder, "colours|Name,Hex");
        foreach (GanttColourToken colour in Colours)
        {
            AppendLine(builder, Join(colour.Name, colour.HexValue));
        }

        AppendLine(builder, "typography|Name,Value,IsBold");
        foreach (GanttTypographyToken typography in Typography)
        {
            AppendLine(builder, Join(typography.Name, typography.Value, typography.IsBold ? "TRUE" : "FALSE"));
        }

        AppendTableHeader(builder, SettingsTableName, SettingsHeaders);
        foreach (GanttSettingDefinition setting in Settings)
        {
            AppendLine(builder, Join(setting.Key, setting.DefaultValue));
        }

        var digest = System.Security.Cryptography.SHA256.HashData(
            System.Text.Encoding.UTF8.GetBytes(builder.ToString()));
        return Convert.ToHexStringLower(digest);
    }

    private static void AppendTableHeader(System.Text.StringBuilder builder, string tableName, string[] headers) =>
        AppendLine(builder, $"{tableName}|{string.Join(",", headers)}");

    private static void AppendLine(System.Text.StringBuilder builder, string line)
    {
        _ = builder.Append(line);
        _ = builder.Append('\n');
    }

    private static string Join(params string[] values) => string.Join("|", values);

    private static string Format(double value) => value.ToString("R", CultureInfo.InvariantCulture);
}
