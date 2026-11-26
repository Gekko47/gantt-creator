namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Tripwire: a test class that can reach the process-wide diagnostics singletons must join the
/// <c>diagnostics-service</c> collection.
/// </summary>
/// <remarks>
/// <para>
/// The defect of 2026-11-24, traced from a real failure: <c>RibbonValueContractTests</c> ran in
/// parallel with <c>AddInHostTests</c> (xUnit runs different collections concurrently), clicked a
/// plot toggle, and <c>RibbonStateService.TryWriteSettings</c> wrote
/// <c>gantt.settings persist refused keys=4 …</c> through <c>CommandBoundary.Instance</c> — the
/// process-wide boundary whose log reference the other test had just installed. The record landed
/// between that test's <c>open</c> and <c>close</c> lines, so its exact-sequence assertion failed
/// intermittently.
/// </para>
/// <para>
/// The stack that named it:
/// <c>TryWriteSettings ← TryPersistPlotRange ← SetPlotStartAuto ← OnPlotStartAutoClick</c>, from
/// <c>RibbonValueContractTests.The_auto_checkboxes_read_back_exactly_what_they_set</c>.
/// </para>
/// <para>
/// Membership is the fix, not luck: xUnit serialises the tests of one collection, so nothing in
/// the collection can inject a record while another member is asserting on the log.
/// </para>
/// </remarks>
public class RibbonTestCollectionGuardTests
{
    /// <summary>Source markers that mean the file can reach the diagnostics singletons.</summary>
    private static readonly string[] SingletonMarkers =
    [
        "RibbonStateService",
        "GanttRibbon.",
        "DiagnosticsService.Instance",
        "CommandBoundary.Instance",
    ];

    /// <summary>Source markers that mean the file can run a test.</summary>
    /// <remarks>
    /// Only a file that declares a test can interleave a record into another test's log, so a
    /// shared fixture or an abstract base is out of scope however many singletons it names. The
    /// two markers are the attributes xUnit needs; a file carrying neither cannot run.
    /// </remarks>
    private static readonly string[] TestMarkers =
    [
        "[Fact",
        "[Theory",
    ];

    /// <summary>The collection every such file must join.</summary>
    private const string CollectionName = "diagnostics-service";

    /// <summary>This file is the rule, so it is not subject to it.</summary>
    private const string ThisFileName = "RibbonTestCollectionGuardTests.cs";

    /// <summary>
    /// Fails listing every test file that reaches a diagnostics singleton without joining the
    /// serialised collection.
    /// </summary>
    [Fact]
    public void Every_test_file_that_reaches_the_diagnostics_singletons_joins_the_collection()
    {
        string directory = LocateTestSourceDirectory();
        var offenders = Directory
            .EnumerateFiles(directory, "*.cs", SearchOption.TopDirectoryOnly)
            .Where(path => !string.Equals(Path.GetFileName(path), ThisFileName, StringComparison.Ordinal))
            .Where(path => Violates(File.ReadAllText(path)))
            .Select(path => Path.GetFileName(path))
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToArray();

        Assert.True(
            offenders.Length == 0,
            "Test files that reach the diagnostics singletons outside the serialised collection "
                + "(add [Collection(\"" + CollectionName + "\")]): " + string.Join(", ", offenders));
    }

    /// <summary>
    /// The positive control: the classifier must fire on a hazard and stay silent on a compliant
    /// file, a shared fixture and an unrelated one, or the tripwire above could pass for the wrong
    /// reason.
    /// </summary>
    [Fact]
    public void The_classifier_flags_a_hazard_and_clears_a_compliant_file()
    {
        const string hazard =
            "[Fact] public class X { void M() { GanttRibbon.OnPlotStartAutoClick(null!, false, service); } }";
        const string compliant =
            "[Collection(\"diagnostics-service\")] [Fact] public class X { void M() { new RibbonStateService(); } }";
        const string sharedFixture =
            "internal abstract class X { protected static RibbonStateService Arm() => new(); }";
        const string unrelated = "[Fact] public class X { void M() { _ = 1 + 1; } }";

        Assert.True(Violates(hazard), "a parallel test class driving a ribbon callback must be flagged");
        Assert.False(Violates(compliant), "a member of the collection must not be flagged");
        Assert.False(Violates(sharedFixture), "a file that cannot run a test must not be flagged");
        Assert.False(Violates(unrelated), "an unrelated test file must not be flagged");
    }

    /// <summary>Whether one test file breaks the rule.</summary>
    /// <param name="source">The file's text.</param>
    /// <returns><see langword="true"/> when it is a hazard outside the collection.</returns>
    private static bool Violates(string source) =>
        DeclaresTests(source)
        && ReachesSingletons(source)
        && !source.Contains(CollectionName, StringComparison.Ordinal);

    /// <summary>Whether the source declares at least one test.</summary>
    /// <param name="source">The file's text.</param>
    /// <returns><see langword="true"/> when the file can run a test.</returns>
    private static bool DeclaresTests(string source) =>
        TestMarkers.Any(marker => source.Contains(marker, StringComparison.Ordinal));

    /// <summary>Whether the source touches any diagnostics singleton or ribbon callback surface.</summary>
    /// <param name="source">The file's text.</param>
    /// <returns><see langword="true"/> when it can reach a singleton.</returns>
    private static bool ReachesSingletons(string source) =>
        SingletonMarkers.Any(marker => source.Contains(marker, StringComparison.Ordinal));

    /// <summary>
    /// Finds <c>tests/GanttCreator.AddIn.Tests</c> by walking up from the test binary, so the same
    /// scan runs locally and on a hosted runner regardless of where the repository is checked out.
    /// </summary>
    /// <returns>The directory holding this test project's source files.</returns>
    /// <exception cref="InvalidOperationException">The source directory could not be located.</exception>
    private static string LocateTestSourceDirectory()
    {
        DirectoryInfo? current = new(AppContext.BaseDirectory);
        while (current is not null)
        {
            string candidate = Path.Combine(current.FullName, "tests", "GanttCreator.AddIn.Tests");
            if (Directory.Exists(candidate))
            {
                return candidate;
            }

            current = current.Parent;
        }

        throw new InvalidOperationException(
            "RibbonTestCollectionGuardTests could not locate tests/GanttCreator.AddIn.Tests above "
                + AppContext.BaseDirectory);
    }
}
