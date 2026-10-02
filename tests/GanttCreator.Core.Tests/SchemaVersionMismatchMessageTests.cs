using GanttCreator.Core;
using GanttCreator.Core.ConfigIntegrity;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="SchemaVersionMismatchMessage"/>, the Ribbon-visible detail
/// R4.7C D6 requires for a schema mismatch.
/// </summary>
/// <remarks>
/// <para>
/// The validator ships with its positive test in the same change: a version-3
/// workbook is <em>reported</em> rather than converted (ADR-0029 D6), so the only
/// thing standing between a user and a silent data problem is this string.
/// </para>
/// </remarks>
public sealed class SchemaVersionMismatchMessageTests
{
    [Fact]
    public void The_message_names_both_versions_and_the_remedy()
    {
        string message = SchemaVersionMismatchMessage.Build(3, 5);

        Assert.Contains("3", message, StringComparison.Ordinal);
        Assert.Contains("5", message, StringComparison.Ordinal);

        // Actionable, not a generic integrity string: it must name the remedy.
        Assert.Contains("no automatic upgrade", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Initialise Gantt Sheet", message, StringComparison.Ordinal);
    }

    [Fact]
    public void A_newer_workbook_is_reported_with_the_same_remedy()
    {
        // A workbook ahead of the add-in is a different situation from an old one,
        // but the refusal is the same and the remedy is the same: the add-in
        // cannot read a schema it does not implement. The message must not imply
        // the workbook is simply out of date in both directions.
        string message = SchemaVersionMismatchMessage.Build(6, 5);

        Assert.Contains("6", message, StringComparison.Ordinal);
        Assert.Contains("5", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(-1, 5)]
    [InlineData(5, 0)]
    [InlineData(5, -1)]
    public void A_non_positive_version_is_refused(int workbookVersion, int expectedVersion)
    {
        // Positive test for the validator: a non-positive version is not a version
        // this builder can describe, and rendering "version 0" would be less honest
        // than refusing.
        Assert.Throws<ArgumentOutOfRangeException>(
            () => SchemaVersionMismatchMessage.Build(workbookVersion, expectedVersion));
    }

    [Fact]
    public void The_current_version_builds_a_message_against_itself()
    {
        // Not an error path, but it proves the builder is usable for the version the
        // add-in actually ships, and that the numbers come from the constant rather
        // than from a literal pinned beside it.
        string message = SchemaVersionMismatchMessage.Build(
            GanttSchemaVersion.CurrentSchemaVersion,
            GanttSchemaVersion.CurrentSchemaVersion);

        Assert.Contains(
            GanttSchemaVersion.CurrentSchemaVersion.ToString(System.Globalization.CultureInfo.InvariantCulture),
            message,
            StringComparison.Ordinal);
    }
}
