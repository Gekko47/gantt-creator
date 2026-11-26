using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Office;
using Moq;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The fixtures the ribbon-state test classes share.
/// </summary>
/// <remarks>
/// Each class keeps its own constructor and <see cref="IDisposable"/> reset, because the session
/// singleton must be cleared around every test and an inherited lifecycle hides that. What is
/// shared is only the fixture construction, reached through <c>using static</c> so the call sites
/// stay unchanged. No test lives here.
/// </remarks>
internal static class RibbonStateTestSupport
{
    /// <summary>Arms the session singleton over the supplied adapter and ribbon.</summary>
    /// <param name="adapter">The Excel application substitute.</param>
    /// <param name="logSource">The log-availability source, or null.</param>
    /// <param name="ribbon">The ribbon substitute, or null.</param>
    /// <returns>The armed session service.</returns>
    internal static RibbonStateService Arm(
        TestExcelApplicationAdapter adapter,
        Func<bool?>? logSource = null,
        IRibbonUI? ribbon = null)
    {
        var service = RibbonStateService.Instance;
        service.SetApplicationAdapter(adapter);
        service.SetLogAvailabilitySource(logSource);
        service.SetRibbon(ribbon);
        service.Activate();
        return service;
    }

    /// <summary>Counts the full ribbon invalidations the service requested.</summary>
    /// <param name="ribbon">The ribbon substitute.</param>
    /// <returns>The number of <see cref="IRibbonUI.Invalidate"/> calls.</returns>
    internal static int CountInvalidations(Mock<IRibbonUI> ribbon)
        => ribbon.Invocations.Count(invocation => invocation.Method.Name == nameof(IRibbonUI.Invalidate));

    /// <summary>
    /// A well-formed <see cref="GanttRowId"/> for seeded table rows:
    /// <c>G-</c> plus 32 lowercase hex characters.
    /// </summary>
    /// <param name="last">The final character, so two rows differ.</param>
    /// <returns>The row identity.</returns>
    internal static string NewRowId(char last = 'a') => "G-" + new string('0', 31) + last;

    /// <summary>Arms the session and seeds it with one table-read outcome.</summary>
    /// <param name="adapter">The Excel application substitute.</param>
    /// <param name="outcome">The outcome the table reader returns.</param>
    /// <returns>The refreshed session service.</returns>
    internal static RibbonStateService ArmedWithTableRead(
        TestExcelApplicationAdapter adapter,
        GanttTableReadOutcome outcome)
    {
        var service = Arm(adapter);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read()).Returns(outcome);
        service.SetTableReader(reader.Object);
        service.Refresh();
        return service;
    }

    /// <summary>One table body row of the supplied type.</summary>
    /// <param name="rowNumber">The row number.</param>
    /// <param name="type">The entity type.</param>
    /// <param name="parentId">The parent row identity, or null for a top-level row.</param>
    /// <returns>The row.</returns>
    internal static GanttRowDto Row(int rowNumber, string type, string? parentId = null) =>
        new(
            rowNumber,
            NewRowId((char)('a' + rowNumber)),
            null,
            0,
            type,
            null,
            null,
            null,
            parentId,
            null,
            null,
            null,
            null,
            true,
            null);
}
