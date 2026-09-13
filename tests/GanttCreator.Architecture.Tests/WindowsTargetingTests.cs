using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Enforces that every project targeting a <c>-windows</c> TFM declares
/// <c>EnableWindowsTargeting=true</c>. On a non-Windows host the .NET SDK
/// cannot resolve the Windows targeting pack otherwise, and restore fails
/// with NETSDK1100 (observed when the locked-mode restore step ran on a
/// Linux runner); the property is a no-op on Windows builds. Mirrors the
/// docs/08-TEST-CHECKLIST.md build-pipeline traceability contract: the
/// discovered set is pinned to the explicit list below so the enumeration
/// cannot silently pass empty or drift.
/// </summary>
public sealed class WindowsTargetingTests
{
    /// <summary>
    /// Projects expected to declare a <c>-windows</c> target framework.
    /// Keep in sync with
    /// <see cref="Enable_windows_targeting_on_every_windows_tfm_project"/>.
    /// </summary>
    private static readonly string[] ExpectedWindowsProjects =
    [
        "src/GanttCreator.AddIn/GanttCreator.AddIn.csproj",
        "src/GanttCreator.Office/GanttCreator.Office.csproj",
        "src/GanttCreator.Raster/GanttCreator.Raster.csproj",
        "tests/GanttCreator.AddIn.Tests/GanttCreator.AddIn.Tests.csproj",
        "tests/GanttCreator.Office.ContractTests/GanttCreator.Office.ContractTests.csproj",
        "tests/GanttCreator.Office.IntegrationTests/GanttCreator.Office.IntegrationTests.csproj",
        "tests/GanttCreator.Raster.Tests/GanttCreator.Raster.Tests.csproj",
    ];

    [Fact]
    public void Enable_windows_targeting_on_every_windows_tfm_project()
    {
        var discovered = EnumerateWindowsProjects().ToArray();

        // A definitely-non-empty discovery is a hard requirement: an empty
        // enumeration must fail here rather than let the loop below pass.
        Assert.NotEmpty(discovered);

        Assert.Equal(
            ExpectedWindowsProjects.Order(StringComparer.Ordinal),
            discovered.Select(p => p.RelativePath).Order(StringComparer.Ordinal));

        foreach (var project in discovered)
        {
            var doc = XDocument.Load(project.AbsolutePath);
            XNamespace ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var enable = doc.Descendants(ns + "EnableWindowsTargeting").Select(e => e.Value.Trim()).ToArray();

            Assert.True(
                enable.Length == 1,
                $"Project {project.RelativePath} must declare EnableWindowsTargeting exactly once; found {enable.Length} occurrence(s).");
            Assert.Equal("true", enable[0], StringComparer.Ordinal);
        }
    }

    private sealed record WindowsProject(string RelativePath, string AbsolutePath);

    private static IEnumerable<WindowsProject> EnumerateWindowsProjects()
    {
        var root = FindRepoRoot();

        foreach (var path in Directory.EnumerateFiles(root, "*.csproj", SearchOption.AllDirectories))
        {
            var doc = XDocument.Load(path);
            XNamespace ns = doc.Root?.GetDefaultNamespace() ?? XNamespace.None;
            var tfm = doc.Descendants(ns + "TargetFramework").FirstOrDefault()?.Value
                ?? doc.Descendants(ns + "TargetFrameworks").FirstOrDefault()?.Value
                ?? string.Empty;

            if (tfm.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Any(t => t.EndsWith("-windows", StringComparison.Ordinal)))
            {
                yield return new WindowsProject(
                    Path.GetRelativePath(root, path).Replace('\\', '/'),
                    path);
            }
        }
    }

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(WindowsTargetingTests).Assembly.Location)!);
        while (dir is not null)
        {
            if (File.Exists(Path.Combine(dir.FullName, "GanttCreator.slnx"))
                && Directory.Exists(Path.Combine(dir.FullName, "src"))
                && Directory.Exists(Path.Combine(dir.FullName, "tests")))
            {
                return dir.FullName;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate the repository root from the test binary.");
    }
}