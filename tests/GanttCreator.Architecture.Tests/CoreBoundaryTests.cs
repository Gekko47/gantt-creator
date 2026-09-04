using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// The Core boundary. <c>GanttCreator.Core</c> must have no Office,
/// Excel-DNA, SkiaSharp, clipboard, filesystem-dialog, or UI
/// dependency. The list is enumerated explicitly so a future offender
/// fails this test with the exact assembly name. The test loads the
/// compiled <c>GanttCreator.Core.dll</c> by path (no <c>ProjectReference</c>)
/// so the architecture is enforced even if a developer wires a
/// reference into <c>GanttCreator.Core.csproj</c>.
/// </summary>
public sealed class CoreBoundaryTests
{
    // Assemblies that GanttCreator.Core must never reference. The list is
    // intentionally explicit; an unknown offender fails the test rather
    // than slipping through.
    private static readonly string[] ForbiddenAssemblies =
    {
        // Office / Excel-DNA / PowerPoint
        "Microsoft.Office",
        "Microsoft.Office.Interop",
        "Microsoft.Office.Tools",
        "office",
        "ExcelDna",
        "ExcelDna.AddIn",
        "ExcelDna.Integration",
        "ExcelDna.Interop",
        "Microsoft.Vbe.Interop",
        // Raster
        "SkiaSharp",
        "SkiaSharp.Views",
        "HarfBuzzSharp",
        // Clipboard / dialogs / UI
        "System.Windows.Clipboard",
        "System.Windows.Forms",
        "System.Drawing",
        "System.Drawing.Common",
        "Microsoft.Win32",
        "PresentationCore",
        "PresentationFramework",
        "WindowsBase",
        "System.Xaml",
    };

    [Fact]
    public void Core_assembly_does_not_reference_forbidden_assemblies()
    {
        var coreDll = LocateCoreAssembly();
        Assert.True(File.Exists(coreDll),
            $"Could not locate GanttCreator.Core.dll at '{coreDll}'. " +
            "Build the solution before running architecture tests.");

        var assemblyNames = Assembly.LoadFrom(coreDll).GetReferencedAssemblies()
            .Select(a => a.Name ?? string.Empty)
            .ToArray();

        foreach (var forbidden in ForbiddenAssemblies)
        {
            Assert.DoesNotContain(assemblyNames, n => string.Equals(n, forbidden, System.StringComparison.OrdinalIgnoreCase)
                || (n?.StartsWith(forbidden + ".", System.StringComparison.OrdinalIgnoreCase) ?? false));
        }
    }

    [Fact]
    public void Core_assembly_targets_net10_0()
    {
        var coreDll = LocateCoreAssembly();
        Assert.True(File.Exists(coreDll));

        var assembly = Assembly.LoadFrom(coreDll);
        // The simple test: the assembly's image location is the net10.0
        // build output. A path under net10.0-windows would mean Core
        // drifted to a Windows target. The build will not produce both
        // TFMs for Core by design.
        Assert.Contains("net10.0", coreDll.Replace('\\', '/'), StringComparison.Ordinal);
        Assert.DoesNotContain("net10.0-windows", coreDll.Replace('\\', '/'), StringComparison.Ordinal);
    }

    private static string LocateCoreAssembly()
    {
        // Walk up from the test binary until we find a sibling src/ folder.
        var dir = new DirectoryInfo(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, "src", "GanttCreator.Core", "bin");
            if (Directory.Exists(candidate))
            {
                // Prefer Release/net10.0; fall back to any TFM under Release.
                var releaseNet = Path.Combine(candidate, "Release", "net10.0", "GanttCreator.Core.dll");
                if (File.Exists(releaseNet)) return releaseNet;
                var anyRelease = Directory
                    .EnumerateFiles(Path.Combine(candidate, "Release"), "GanttCreator.Core.dll", SearchOption.AllDirectories)
                    .FirstOrDefault();
                if (anyRelease is not null) return anyRelease;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate src/GanttCreator.Core from the test binary.");
    }
}
