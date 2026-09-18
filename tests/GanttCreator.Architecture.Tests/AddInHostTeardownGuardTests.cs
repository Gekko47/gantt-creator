using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Pins the work-item R1.6 D1 teardown contract in
/// <c>src/GanttCreator.AddIn/AddInHost.cs</c>: every teardown operation in
/// <c>AutoClose</c> — the owned COM detach, the close record, and both
/// log-reference drops — is enclosed by its own catch clause, so one failing
/// step cannot skip the steps after it (specifically,
/// <c>DiagnosticsService.Reset()</c> still runs when
/// <c>CommandBoundary.Reset()</c> fails).
/// </summary>
/// <remarks>
/// The three <c>Reset</c> calls are non-throwing by contract, so no runtime
/// test can make one fail; the protecting shape is asserted on the parsed
/// source instead, with a positive control for the pre-fix shared-guard
/// shape. The scan is deliberately shape-sensitive: a refactor that moves
/// teardown behind a shared helper must be re-reviewed here rather than
/// passing silently.
/// </remarks>
public sealed class AddInHostTeardownGuardTests
{
    private const string AddInHostRelativePath = "src/GanttCreator.AddIn/AddInHost.cs";

    /// <summary>
    /// The teardown operations that must each have their own guard. Matched
    /// against the invocation's callee text, so comments cannot satisfy the
    /// check.
    /// </summary>
    private static readonly string[] TeardownOperations =
    [
        "RibbonStateService.Reset",
        "LogClose",
        "CommandBoundary.Reset",
        "DiagnosticsService.Reset",
    ];

    [Fact]
    public void AutoClose_guards_each_teardown_operation_separately()
    {
        var source = File.ReadAllText(LocateRepoFile(AddInHostRelativePath));

        var violations = FindSharedGuardViolations(source);

        Assert.True(
            violations.Count == 0,
            "AutoClose teardown guards: " + string.Join(" | ", violations));
    }

    // __POS__
    #region Positive controls

    [Fact]
    public void Checker_rejects_the_shared_guard_shape()
    {
        const string SharedGuardShape = """
            class Host
            {
                void AutoClose()
                {
                    try
                    {
                        RibbonStateService.Reset();
                        try
                        {
                            lifecycle.LogClose(token);
                        }
                        catch
                        {
                        }

                        CommandBoundary.Reset();
                        DiagnosticsService.Reset();
                    }
                    catch
                    {
                    }
                }
            }
            """;

        Assert.NotEmpty(FindSharedGuardViolations(SharedGuardShape));
    }

    [Fact]
    public void Checker_accepts_the_per_step_guard_shape()
    {
        const string PerStepGuardShape = """
            class Host
            {
                void AutoClose()
                {
                    try { RibbonStateService.Reset(); } catch { }
                    try { lifecycle.LogClose(token); } catch { }
                    try { CommandBoundary.Reset(); } catch { }
                    try { DiagnosticsService.Reset(); } catch { }
                    try { log?.Dispose(); } catch { } finally { log = null; }
                }
            }
            """;

        Assert.Empty(FindSharedGuardViolations(PerStepGuardShape));
    }

    [Fact]
    public void Checker_rejects_a_missing_teardown_operation()
    {
        // A deleted step must fail the contract instead of passing vacuously.
        const string MissingStepShape = """
            class Host
            {
                void AutoClose()
                {
                    try { RibbonStateService.Reset(); } catch { }
                    try { lifecycle.LogClose(token); } catch { }
                    try { CommandBoundary.Reset(); } catch { }
                }
            }
            """;

        Assert.Contains(
            FindSharedGuardViolations(MissingStepShape),
            violation => violation.Contains("DiagnosticsService.Reset", StringComparison.Ordinal));
    }

    #endregion

    // __INFRA__
    #region Infrastructure

    /// <summary>
    /// Reports every teardown operation in <c>AutoClose</c> that is not
    /// enclosed by exactly one catch-bearing <c>try</c>, plus every pair of
    /// operations that share one guard.
    /// </summary>
    /// <param name="source">The C# source to inspect.</param>
    /// <returns>One message per violation; empty when the contract holds.</returns>
    private static List<string> FindSharedGuardViolations(string source)
    {
        var root = CSharpSyntaxTree.ParseText(source).GetRoot();
        var autoClose = root.DescendantNodes()
            .OfType<MethodDeclarationSyntax>()
            .FirstOrDefault(method => string.Equals(method.Identifier.ValueText, "AutoClose", StringComparison.Ordinal));

        if (autoClose is null)
        {
            return ["AutoClose was not found in the inspected source."];
        }

        var tryStatements = autoClose.DescendantNodes().OfType<TryStatementSyntax>().ToList();
        var guards = new Dictionary<string, TryStatementSyntax>(StringComparer.Ordinal);
        var violations = new List<string>();

        foreach (var operation in TeardownOperations)
        {
            var owners = tryStatements.Where(candidate => ContainsOperation(candidate, operation)).ToList();

            if (owners.Count != 1)
            {
                violations.Add(
                    $"{operation} is enclosed by {owners.Count} try statement(s); expected exactly one per-step guard with a catch clause.");
                continue;
            }

            if (owners[0].Catches.Count == 0)
            {
                violations.Add($"{operation} has no catch clause around it.");
                continue;
            }

            guards[operation] = owners[0];
        }

        foreach (var group in guards.GroupBy(pair => pair.Value).Where(group => group.Count() > 1))
        {
            violations.Add(
                "shared guard for: " + string.Join(", ", group.Select(pair => pair.Key).Order(StringComparer.Ordinal)));
        }

        return violations;
    }

    /// <summary>
    /// Returns whether <paramref name="candidate"/> encloses an invocation of
    /// <paramref name="operation"/>. Matching the callee text on the syntax
    /// node excludes comments and string literals.
    /// </summary>
    private static bool ContainsOperation(TryStatementSyntax candidate, string operation) =>
        candidate.DescendantNodes()
            .OfType<InvocationExpressionSyntax>()
            .Any(invocation => MatchesOperation(invocation, operation));

    private static bool MatchesOperation(InvocationExpressionSyntax invocation, string operation)
    {
        var callee = invocation.Expression.ToString();
        return string.Equals(operation, "LogClose", StringComparison.Ordinal)
            ? callee.EndsWith(".LogClose", StringComparison.Ordinal)
            : string.Equals(callee, operation, StringComparison.Ordinal);
    }

    private static string LocateRepoFile(string relativePath)
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(AddInHostTeardownGuardTests).Assembly.Location)!);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "GanttCreator.slnx"))
                && Directory.Exists(Path.Combine(dir.FullName, "src")))
            {
                var candidate = Path.Combine(dir.FullName, relativePath.Replace('/', Path.DirectorySeparatorChar));
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException($"Could not locate '{relativePath}' from the test binary.");
    }

    #endregion
}