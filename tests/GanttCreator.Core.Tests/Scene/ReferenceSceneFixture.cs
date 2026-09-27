using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace GanttCreator.Core.Tests.Scene;

/// <summary>
/// The R3.12 canonical neutral construction-delay fixture, loaded from the one
/// committed <c>tests/fixtures/reference-gantt.json</c>.
/// </summary>
/// <remarks>
/// <para>
/// One input serves the Core scene tests and the Office integration tests, so the
/// representative Gantt is defined once rather than restated per suite. Nothing here
/// is random and no row identifier is generated: every value is fixed in the file so
/// the scene serialization is byte-stable across runs and machines.
/// </para>
/// <para>
/// The fixture is data only. It is normalised through the landed R2 pipeline
/// (<see cref="GanttRowValidator"/>) exactly as a worksheet read would be, so a
/// change in row validation moves this fixture rather than diverging from it.
/// </para>
/// </remarks>
internal static class ReferenceSceneFixture
{
    /// <summary>The inclusive plot start date for the fixture chart.</summary>
    public static readonly DateOnly PlotStart = new(2026, 1, 5);

    /// <summary>The inclusive plot finish date for the fixture chart.</summary>
    public static readonly DateOnly PlotFinish = new(2026, 4, 24);

    /// <summary>The fixture chart title.</summary>
    public const string Title = "Site Enabling Works";

    // The fixture file is authored in camelCase for readability, so the naming
    // policy is mapped explicitly rather than left case-sensitive. Unmapped members
    // are still rejected: an unknown field means the file and this loader have
    // drifted, which must fail loudly rather than silently ignore a row.
    private static readonly JsonSerializerOptions _options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
    };

    /// <summary>Loads the fixture body rows in file order.</summary>
    /// <returns>The raw <c>tblGanttData</c> body rows.</returns>
    public static IReadOnlyList<GanttRowDto> LoadRows()
    {
        List<ReferenceFixtureRow>? rows = LoadDocument().Rows;
        if (rows is null)
        {
            throw new InvalidDataException("The reference Gantt fixture has no rows.");
        }

        return [.. rows.Select(ToRowDto)];
    }

    /// <summary>Loads the fixture and runs it through the landed row validator.</summary>
    /// <returns>The validation outcome, including the renderable events.</returns>
    public static GanttValidationOutcome LoadValidated() => GanttRowValidator.Validate([.. LoadRows()]);

    /// <summary>Gets the path to the committed fixture file.</summary>
    /// <returns>The absolute fixture path.</returns>
    public static string Path_ =>
        System.IO.Path.Combine(AppContext.BaseDirectory, "fixtures", "reference-gantt.json");

    private static ReferenceFixtureDocument LoadDocument()
    {
        string json = File.ReadAllText(Path_);
        ReferenceFixtureDocument document = JsonSerializer.Deserialize<ReferenceFixtureDocument>(json, _options)
            ?? throw new InvalidDataException("The reference Gantt fixture is empty.");
        if (document.SchemaVersion != 1)
        {
            throw new InvalidDataException($"Unsupported reference fixture schema version {document.SchemaVersion}.");
        }

        return document;
    }

    private static GanttRowDto ToRowDto(ReferenceFixtureRow row) =>
        new(
            row.RowNumber,
            row.Id,
            row.LaneId,
            row.StackIndex,
            row.Type,
            row.Description,
            row.Start,
            row.Finish,
            row.ParentId,
            row.StyleKey,
            row.LabelPosition,
            row.FillColour,
            row.StrokeColour,
            null,
            null);
}
