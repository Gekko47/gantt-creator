using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Enforces ADR-0008 D4 — "the workbook-protection guard is the first check in
/// every mutating adapter" — on the Office adapter sources, and keeps the
/// remaining inline-probe debt visible instead of silent.
/// </summary>
/// <remarks>
/// <para>
/// Four rules, each with a positive control:
/// </para>
/// <list type="number">
/// <item>every Office source file is classified either as a data-mutating
/// adapter or as read-only, so a new adapter file cannot be added without a
/// deliberate classification;</item>
/// <item>data-mutation discovery matches the registered adapters, so an
/// unregistered mutating adapter fails;</item>
/// <item>every registered mutating adapter's entry point consults protection,
/// and the file's first protection read precedes its first data
/// mutation;</item>
/// <item>the mutating adapters that still use their own inline protection
/// probes instead of the shared <c>IWorksheetProtectionGuard</c> port are
/// exactly the declared set, and every declared file is named in
/// <c>docs/KNOWN-LIMITATIONS.md</c> (entry L16).</item>
/// </list>
/// <para>
/// The scan is deliberately source-shape based, not semantic: it matches the
/// bulk cell write (<c>.Value2 =</c>, never a <c>==</c> comparison) and the
/// table/sheet/name creation calls this codebase's adapters use
/// (<see cref="DataMutationPatterns"/>). A future adapter that writes through a
/// different COM shape must be re-reviewed here rather than passing silently.
/// Comments are excluded from matching (token text only), so documentation can
/// neither satisfy nor break a rule.
/// </para>
/// <para>
/// Known gap, deliberately recorded rather than hidden:
/// <c>ExcelGanttValidationReporter</c> writes cell notes and has no protection
/// pre-check. It is not a destructive command under ADR-0008 D1 (it never
/// clears or overwrites user cell content or table geometry, and it preserves
/// user note text verbatim), so it is classified read-only here. Adding a
/// protection refusal to it is a separate decision for its own row.
/// </para>
/// </remarks>
public sealed class ProtectionGuardFirstTests
{
    /// <summary>A file that mutates the visible table or the configuration catalogues.</summary>
    /// <param name="RelativePath">Repo-relative path of the adapter source.</param>
    /// <param name="EntryMethod">The method that must consult protection first.</param>
    private sealed record MutatingAdapter(string RelativePath, string EntryMethod);

    /// <summary>
    /// Every Office adapter that writes cell values, table geometry, sheets, or
    /// defined names. Keep in sync with the sources: the discovery rule fails
    /// when a file mutates without being registered here.
    /// </summary>
    private static readonly MutatingAdapter[] MutatingAdapters =
    [
        new("src/GanttCreator.Office/ExcelWorkbookInitialiser.cs", "Initialise"),
        new("src/GanttCreator.Office/ExcelConfigCatalogueWriter.cs", "Write"),
    ];

    /// <summary>
    /// Mutating adapters that still use their own read-only protection probes
    /// (<c>ProtectContents</c>/<c>ProtectStructure</c>) rather than the shared
    /// guard port. ADR-0008 D4 wants the port; this declared set is the
    /// documented debt and must shrink to empty before the R5.8 live
    /// destructive run. Moving a file to the port means deleting it here — the
    /// test fails until the declaration is updated, so the ratchet only
    /// tightens deliberately.
    /// </summary>
    private static readonly string[] PendingGuardPortAdoption =
    [
        "src/GanttCreator.Office/ExcelWorkbookInitialiser.cs",
        "src/GanttCreator.Office/ExcelConfigCatalogueWriter.cs",
    ];

    /// <summary>
    /// The remaining Office sources: ports, outcome types, readers, the
    /// application-state adapter, and the note-writing validation reporter
    /// (see the type remarks for why it is not in
    /// <see cref="MutatingAdapters"/>). An unlisted file fails the
    /// classification rule.
    /// </summary>
    private static readonly string[] NoDataMutationOfficeFiles =
    [
        "src/GanttCreator.Office/ExcelApplicationAdapter.cs",
        "src/GanttCreator.Office/ExcelConfigCatalogueReader.cs",
        "src/GanttCreator.Office/ExcelGanttTableReader.cs",
        "src/GanttCreator.Office/ExcelGanttValidationReporter.cs",
        "src/GanttCreator.Office/ExcelWorksheetProtectionGuard.cs",
        "src/GanttCreator.Office/GanttTableReadOutcome.cs",
        "src/GanttCreator.Office/GanttValidationReportOutcome.cs",
        "src/GanttCreator.Office/IConfigCatalogueReader.cs",
        "src/GanttCreator.Office/IConfigCatalogueWriter.cs",
        "src/GanttCreator.Office/IExcelApplicationAdapter.cs",
        "src/GanttCreator.Office/IGanttTableReader.cs",
        "src/GanttCreator.Office/IGanttValidationReporter.cs",
        "src/GanttCreator.Office/InternalsVisibleTo.cs",
        "src/GanttCreator.Office/IWorkbookInitialiser.cs",
        "src/GanttCreator.Office/IWorksheetProtectionGuard.cs",
        "src/GanttCreator.Office/ProtectionGuardOutcome.cs",
        "src/GanttCreator.Office/ValidationReportComposer.cs",
        "src/GanttCreator.Office/WorkbookInitialiseOutcome.cs",
    ];

    /// <summary>
    /// The shared guard port's type name. A registered adapter either mentions
    /// it (adopted) or is declared in
    /// <see cref="PendingGuardPortAdoption"/> (documented debt).
    /// </summary>
    private const string GuardPortTypeName = "IWorksheetProtectionGuard";

    /// <summary>The evidence ledger that must name every pending-adoption file.</summary>
    private const string LedgerRelativePath = "docs/KNOWN-LIMITATIONS.md";

    private const string OfficeSourceDirectory = "src/GanttCreator.Office";

    /// <summary>
    /// Literal data-mutation shapes. <c>Value2 =</c> is matched as an
    /// assignment (a following <c>=</c> makes it a comparison, not a
    /// mutation); the rest are creation calls the adapters issue against COM
    /// collections.
    /// </summary>
    private static readonly string[] DataMutationPatterns =
    [
        "ListObjects.Add(",
        "listObjects.Add(",
        "ListRows.Add(",
        "listRows.Add(",
        "ListColumns.Add(",
        "listColumns.Add(",
        "Sheets.Add(",
        "sheets.Add(",
        "Names.Add(",
        "names.Add(",
    ];

    /// <summary>
    /// Substrings that mark a read-only protection consultation: the
    /// initialiser's probe seams, the PIA probe properties, and the shared
    /// guard port's single method call. Matched on token text with string
    /// literals blanked, so neither a comment nor a message string can satisfy
    /// the rule, and a refusal-reason name such as <c>TargetProtected</c>
    /// cannot either.
    /// </summary>
    private static readonly string[] ProtectionReadPatterns =
    [
        "IsWorksheetProtected",
        "IsWorkbookStructureProtected",
        "ProtectContents",
        "ProtectStructure",
        ".Query(",
    ];

    [Fact]
    public void Every_office_source_file_is_classified()
    {
        var sources = ReadOfficeSources();

        var violations = ClassificationViolations(
            sources,
            MutatingAdapters.Select(adapter => adapter.RelativePath).ToArray(),
            NoDataMutationOfficeFiles);

        Assert.True(
            violations.Count == 0,
            "Unclassified Office source files: " + string.Join(" | ", violations));
    }

    [Fact]
    public void Data_mutation_discovery_matches_the_registered_adapters()
    {
        var sources = ReadOfficeSources();

        var violations = DiscoveryViolations(
            sources,
            MutatingAdapters.Select(adapter => adapter.RelativePath).ToArray());

        Assert.True(
            violations.Count == 0,
            "Mutating-adapter registry drift: " + string.Join(" | ", violations));
    }

    [Fact]
    public void Registered_mutating_adapters_consult_protection_before_their_first_mutation()
    {
        var sources = ReadOfficeSources();

        var violations = ProtectionOrderViolations(sources, MutatingAdapters);

        Assert.True(
            violations.Count == 0,
            "Protection-order violations: " + string.Join(" | ", violations));
    }

    [Fact]
    public void Guard_port_pending_adoption_is_the_declared_and_documented_set()
    {
        var sources = ReadOfficeSources();

        var inlineOnly = MutatingAdapters
            .Select(adapter => adapter.RelativePath)
            .Where(path => !sources[path].Contains(GuardPortTypeName, StringComparison.Ordinal))
            .Order(StringComparer.Ordinal)
            .ToArray();
        var declared = PendingGuardPortAdoption.Order(StringComparer.Ordinal).ToArray();

        Assert.Equal(declared, inlineOnly);

        // Every declared file must be named in the ledger, so the debt cannot
        // exist only inside this test.
        var ledger = File.ReadAllText(LocateRepoFile(LedgerRelativePath));
        var undocumented = declared
            .Where(path => !ledger.Contains(Path.GetFileName(path), StringComparison.Ordinal))
            .ToArray();

        Assert.True(
            undocumented.Length == 0,
            "Pending guard-port adoption is not recorded in " + LedgerRelativePath + ": " +
                string.Join(", ", undocumented));
    }

    // ---- Positive controls (AGENTS.md: every validator ships with one) ----

    [Fact]
    public void Checker_flags_a_mutation_that_precedes_the_protection_read()
    {
        const string MutationFirst = """
            class Adapter
            {
                void Write()
                {
                    sheet.Range["A1"].Value2 = values;
                    if (config.ProtectContents)
                    {
                        return;
                    }
                }
            }
            """;

        var violations = ProtectionOrderViolations(
            Sources(("src/GanttCreator.Office/Adapter.cs", MutationFirst)),
            [new MutatingAdapter("src/GanttCreator.Office/Adapter.cs", "Write")]);

        Assert.NotEmpty(violations);
    }

    [Fact]
    public void Checker_flags_a_mutating_entry_point_with_no_protection_read()
    {
        const string NoProtectionRead = """
            class Adapter
            {
                void Write()
                {
                    sheet.Range["A1"].Value2 = values;
                }
            }
            """;

        var violations = ProtectionOrderViolations(
            Sources(("src/GanttCreator.Office/Adapter.cs", NoProtectionRead)),
            [new MutatingAdapter("src/GanttCreator.Office/Adapter.cs", "Write")]);

        Assert.NotEmpty(violations);
    }

    [Fact]
    public void Checker_accepts_the_guard_port_shape()
    {
        const string GuardPortShape = """
            class Adapter
            {
                private readonly IWorksheetProtectionGuard _guard;

                void Write()
                {
                    if (_guard.Query() != ProtectionGuardOutcome.NotProtected)
                    {
                        return;
                    }

                    sheet.Range["A1"].Value2 = values;
                }
            }
            """;

        var violations = ProtectionOrderViolations(
            Sources(("src/GanttCreator.Office/Adapter.cs", GuardPortShape)),
            [new MutatingAdapter("src/GanttCreator.Office/Adapter.cs", "Write")]);

        Assert.Empty(violations);
    }

    [Fact]
    public void Checker_ignores_comments_and_message_strings()
    {
        // The only protection-looking text is a comment and a string literal;
        // neither may satisfy the rule (and a '==' comparison is not a write).
        const string CommentOnly = """
            class Adapter
            {
                void Write()
                {
                    // ProtectContents is not checked here.
                    sheet.Range["A1"].Value2 = "set ProtectContents first";
                    if (sheet.ProtectContents == true)
                    {
                        return;
                    }
                }
            }
            """;

        var violations = ProtectionOrderViolations(
            Sources(("src/GanttCreator.Office/Adapter.cs", CommentOnly)),
            [new MutatingAdapter("src/GanttCreator.Office/Adapter.cs", "Write")]);

        Assert.NotEmpty(violations);
    }

    [Fact]
    public void Checker_does_not_flag_a_read_only_adapter_file()
    {
        const string MutatingAdapterShape = """
            class Adapter
            {
                void Write()
                {
                    if (IsWorksheetProtected(sheet))
                    {
                        return;
                    }

                    sheet.Range["A1"].Value2 = values;
                }
            }
            """;
        const string ReadOnlyAdapter = """
            class Reader
            {
                object? Read()
                {
                    return sheet.ProtectContents;
                }
            }
            """;

        var registered = new[] { "src/GanttCreator.Office/Adapter.cs" };

        Assert.Empty(DiscoveryViolations(
            Sources(
                ("src/GanttCreator.Office/Adapter.cs", MutatingAdapterShape),
                ("src/GanttCreator.Office/Reader.cs", ReadOnlyAdapter)),
            registered));
        Assert.Empty(ClassificationViolations(
            Sources(
                ("src/GanttCreator.Office/Adapter.cs", MutatingAdapterShape),
                ("src/GanttCreator.Office/Reader.cs", ReadOnlyAdapter)),
            registered,
            ["src/GanttCreator.Office/Reader.cs"]));
    }

    [Fact]
    public void Checker_flags_an_unregistered_mutating_file()
    {
        const string Unregistered = """
            class Inserter
            {
                void Add()
                {
                    table.ListRows.Add();
                }
            }
            """;

        var violations = DiscoveryViolations(
            Sources(
                ("src/GanttCreator.Office/Adapter.cs", "class Adapter { }"),
                ("src/GanttCreator.Office/Inserter.cs", Unregistered)),
            ["src/GanttCreator.Office/Adapter.cs"]);

        Assert.NotEmpty(violations);
    }

    [Fact]
    public void Checker_flags_a_stale_registration()
    {
        // A registered file that no longer mutates anything is drift too: the
        // registry must not claim protection work for a read-only file.
        var violations = DiscoveryViolations(
            Sources(("src/GanttCreator.Office/Adapter.cs", "class Adapter { }")),
            ["src/GanttCreator.Office/Adapter.cs"]);

        Assert.NotEmpty(violations);
    }

    [Fact]
    public void Checker_reports_an_unclassified_file()
    {
        var violations = ClassificationViolations(
            Sources(
                ("src/GanttCreator.Office/Adapter.cs", "class Adapter { }"),
                ("src/GanttCreator.Office/Newcomer.cs", "class Newcomer { }")),
            ["src/GanttCreator.Office/Adapter.cs"],
            []);

        Assert.NotEmpty(violations);
    }

    // ---- Checker ----

    /// <summary>
    /// Returns one message per mutating-adapter source that is neither
    /// registered nor represented in the declared read-only set.
    /// </summary>
    /// <param name="sources">Repo-relative path → source text.</param>
    /// <param name="registeredPaths">The registered mutating adapters.</param>
    /// <param name="readOnlyPaths">The declared read-only Office sources.</param>
    /// <returns>One message per unclassified file; empty when every file is classified.</returns>
    private static List<string> ClassificationViolations(
        Dictionary<string, string> sources,
        string[] registeredPaths,
        string[] readOnlyPaths)
    {
        var classified = new HashSet<string>(registeredPaths, StringComparer.Ordinal);
        foreach (var path in readOnlyPaths)
        {
            _ = classified.Add(path);
        }

        return sources.Keys
            .Where(path => !classified.Contains(path))
            .Order(StringComparer.Ordinal)
            .Select(path => $"{path} is not classified as a mutating adapter or as read-only.")
            .ToList();
    }

    /// <summary>
    /// Returns one message per discovery mismatch: a file that mutates without
    /// being registered, or a registered file that no longer mutates.
    /// </summary>
    /// <param name="sources">Repo-relative path → source text.</param>
    /// <param name="registeredPaths">The registered mutating adapters.</param>
    /// <returns>One message per mismatch; empty when discovery and registry agree.</returns>
    private static List<string> DiscoveryViolations(
        Dictionary<string, string> sources,
        string[] registeredPaths)
    {
        var discovered = sources
            .Where(pair => ContainsMutationShape(CodeOnly(pair.Value)))
            .Select(pair => pair.Key)
            .Order(StringComparer.Ordinal)
            .ToArray();
        var registered = registeredPaths.Order(StringComparer.Ordinal).ToArray();

        var violations = new List<string>();
        violations.AddRange(discovered
            .Except(registered, StringComparer.Ordinal)
            .Select(path => $"{path} mutates workbook content but is not registered as a mutating adapter."));
        violations.AddRange(registered
            .Except(discovered, StringComparer.Ordinal)
            .Select(path => $"{path} is registered as a mutating adapter but no mutation shape was found in it."));
        return violations;
    }

    /// <summary>
    /// Returns one message per registered adapter whose entry point does not
    /// consult protection, or whose file consults protection only after its
    /// first data mutation.
    /// </summary>
    /// <param name="sources">Repo-relative path → source text.</param>
    /// <param name="adapters">The registered mutating adapters.</param>
    /// <returns>One message per violation; empty when the order rule holds.</returns>
    private static List<string> ProtectionOrderViolations(
        Dictionary<string, string> sources,
        IEnumerable<MutatingAdapter> adapters)
    {
        var violations = new List<string>();

        foreach (var adapter in adapters)
        {
            if (!sources.TryGetValue(adapter.RelativePath, out var source))
            {
                violations.Add($"{adapter.RelativePath} was not found among the Office sources.");
                continue;
            }

            var code = CodeOnly(source);

            if (!ContainsMutationShape(code))
            {
                violations.Add(
                    $"{adapter.RelativePath} is registered as a mutating adapter but no mutation shape was found in it.");
                continue;
            }

            var protectionIndex = FirstPatternIndex(code, ProtectionReadPatterns);
            if (protectionIndex < 0)
            {
                violations.Add(
                    $"{adapter.RelativePath} never consults protection " +
                    "(" + string.Join(", ", ProtectionReadPatterns) + ").");
                continue;
            }

            if (protectionIndex > FirstMutationIndex(code))
            {
                violations.Add(
                    $"{adapter.RelativePath} mutates workbook content before its first protection read.");
            }

            if (!EntryMethodConsultsProtection(source, adapter.EntryMethod))
            {
                violations.Add(
                    $"{adapter.EntryMethod} in {adapter.RelativePath} does not consult protection before it runs.");
            }
        }

        return violations;
    }

    /// <summary>
    /// Returns whether the named method's own tokens include a protection
    /// read, so an adapter cannot keep the check in an unreachable helper.
    /// </summary>
    /// <param name="source">The adapter source text.</param>
    /// <param name="methodName">The entry-point method name.</param>
    /// <returns><see langword="true"/> when the method consults protection.</returns>
    private static bool EntryMethodConsultsProtection(string source, string methodName)
    {
        var method = CSharpSyntaxTree.ParseText(source)
            .GetRoot()
            .DescendantNodes()
            .OfType<MethodDeclarationSyntax>()
            .FirstOrDefault(candidate => string.Equals(
                candidate.Identifier.ValueText,
                methodName,
                StringComparison.Ordinal));

        return method is not null
            && FirstPatternIndex(TokenText(method), ProtectionReadPatterns) >= 0;
    }

    /// <summary>
    /// Returns whether the code contains a data-mutation shape.
    /// </summary>
    /// <param name="code">Token text with comments and string contents removed.</param>
    /// <returns><see langword="true"/> when a mutation shape is present.</returns>
    private static bool ContainsMutationShape(string code) =>
        ContainsAssignment(code, ".Value2")
        || DataMutationPatterns.Any(pattern => code.Contains(pattern, StringComparison.Ordinal));

    /// <summary>
    /// Returns whether the token text assigns to <paramref name="member"/> (a
    /// following <c>=</c> means the occurrence is a comparison, not a write).
    /// </summary>
    /// <param name="code">Token text with comments and string contents removed.</param>
    /// <param name="member">The member name, for example <c>.Value2</c>.</param>
    /// <returns><see langword="true"/> when an assignment is present.</returns>
    private static bool ContainsAssignment(string code, string member)
    {
        var needle = member + "=";
        var index = code.IndexOf(needle, StringComparison.Ordinal);
        while (index >= 0)
        {
            var after = index + needle.Length;
            if (after >= code.Length || code[after] != '=')
            {
                return true;
            }

            index = code.IndexOf(needle, after, StringComparison.Ordinal);
        }

        return false;
    }

    /// <summary>
    /// Returns the earliest index of any pattern, or -1 when none is present.
    /// </summary>
    /// <param name="code">Token text with comments and string contents removed.</param>
    /// <param name="patterns">Ordinal substrings to search for.</param>
    /// <returns>The earliest match index, or -1.</returns>
    private static int FirstPatternIndex(string code, string[] patterns)
    {
        var earliest = -1;
        foreach (var pattern in patterns)
        {
            var index = code.IndexOf(pattern, StringComparison.Ordinal);
            if (index >= 0 && (earliest < 0 || index < earliest))
            {
                earliest = index;
            }
        }

        return earliest;
    }

    /// <summary>
    /// Returns the index of the first data-mutation shape in the token text, or
    /// <see cref="int.MaxValue"/> when there is none.
    /// </summary>
    /// <param name="code">Token text with comments and string contents removed.</param>
    /// <returns>The index of the first mutation, or <see cref="int.MaxValue"/>.</returns>
    private static int FirstMutationIndex(string code)
    {
        var earliest = code.IndexOf(".Value2=", StringComparison.Ordinal);
        if (earliest < 0)
        {
            earliest = int.MaxValue;
        }

        foreach (var pattern in DataMutationPatterns)
        {
            var index = code.IndexOf(pattern, StringComparison.Ordinal);
            if (index >= 0 && index < earliest)
            {
                earliest = index;
            }
        }

        return earliest;
    }

    /// <summary>
    /// Strips comments, string contents, and whitespace from a source file:
    /// the remaining concatenated token text is what every rule matches
    /// against, so documentation and message text cannot influence a verdict.
    /// </summary>
    /// <param name="source">The C# source text.</param>
    /// <returns>The concatenated token text.</returns>
    private static string CodeOnly(string source) =>
        TokenText(CSharpSyntaxTree.ParseText(source).GetRoot());

    /// <summary>
    /// Concatenates the token text of a syntax node, replacing literal tokens
    /// with a placeholder so their contents are ignored.
    /// </summary>
    /// <param name="node">The node to flatten.</param>
    /// <returns>The concatenated token text.</returns>
    private static string TokenText(SyntaxNode node)
    {
        var builder = new StringBuilder();
        foreach (var token in node.DescendantTokens())
        {
            builder.Append(
                token.IsKind(SyntaxKind.StringLiteralToken)
                    || token.IsKind(SyntaxKind.InterpolatedStringTextToken)
                    || token.IsKind(SyntaxKind.CharacterLiteralToken)
                    ? "\"literal\""
                    : token.Text);
        }

        return builder.ToString();
    }

    /// <summary>
    /// Builds the checker's file-set shape from inline test sources.
    /// </summary>
    /// <param name="files">Repo-relative path and source pairs.</param>
    /// <returns>Repo-relative path → source text.</returns>
    private static Dictionary<string, string> Sources(params (string Path, string Source)[] files)
    {
        var sources = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (path, source) in files)
        {
            sources[path] = source;
        }

        return sources;
    }

    /// <summary>
    /// Reads every Office adapter source, keyed by its repo-relative path with
    /// forward slashes.
    /// </summary>
    /// <returns>Repo-relative path → source text.</returns>
    private static Dictionary<string, string> ReadOfficeSources()
    {
        var directory = Path.Combine(LocateRepoRoot(), "src", "GanttCreator.Office");
        var sources = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var file in Directory.GetFiles(directory, "*.cs"))
        {
            sources[OfficeSourceDirectory + "/" + Path.GetFileName(file)] = File.ReadAllText(file);
        }

        return sources;
    }

    /// <summary>
    /// Walks up from the test binary until the repository root (the directory
    /// holding <c>GanttCreator.slnx</c>).
    /// </summary>
    /// <returns>The absolute repository root path.</returns>
    private static string LocateRepoRoot()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(ProtectionGuardFirstTests).Assembly.Location)!);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "GanttCreator.slnx"))
                && Directory.Exists(Path.Combine(dir.FullName, "src")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the repository root from the test binary.");
    }

    /// <summary>
    /// Locates a repository file by walking up from the test binary until the
    /// repository root (the directory holding <c>GanttCreator.slnx</c>).
    /// </summary>
    /// <param name="relativePath">Repo-relative path with forward slashes.</param>
    /// <returns>The absolute path.</returns>
    private static string LocateRepoFile(string relativePath)
    {
        var candidate = Path.Combine(
            LocateRepoRoot(),
            relativePath.Replace('/', Path.DirectorySeparatorChar));
        if (File.Exists(candidate))
        {
            return candidate;
        }

        throw new DirectoryNotFoundException($"Could not locate '{relativePath}' from the test binary.");
    }
}
