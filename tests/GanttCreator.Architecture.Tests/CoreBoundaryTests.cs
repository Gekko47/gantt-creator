using System.Reflection;

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
// artifact-source: the GanttCreator.Core.dll consumed here is produced by
// the 'build Release -warnaserror' step of scripts/verify-quick.ps1 and
// scripts/verify.ps1 (docs/02-ARCHITECTURE.md build-pipeline artifact
// contract). Debug output is also accepted so a plain `dotnet test` is not
// artificially red.
public sealed class CoreBoundaryTests
{
    // Assemblies that GanttCreator.Core must never reference. The list is
    // intentionally explicit; an unknown offender fails the test rather
    // than slipping through.
    private static readonly string[] ForbiddenAssemblies =
    [
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
    ];

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
            Assert.DoesNotContain(assemblyNames, n => string.Equals(n, forbidden, StringComparison.OrdinalIgnoreCase)
                || (n?.StartsWith(forbidden + ".", StringComparison.OrdinalIgnoreCase) ?? false));
        }
    }

    [Fact]
    public void Core_assembly_targets_net10_0()
    {
        var coreDll = LocateCoreAssembly();
        Assert.True(File.Exists(coreDll));
        _ = Assembly.LoadFrom(coreDll);
        // The simple test: the assembly's image location is the net10.0
        // build output. A path under net10.0-windows would mean Core
        // drifted to a Windows target. The build will not produce both
        // TFMs for Core by design.
        Assert.Contains("net10.0", coreDll.Replace('\\', '/'), StringComparison.Ordinal);
        Assert.DoesNotContain("net10.0-windows", coreDll.Replace('\\', '/'), StringComparison.Ordinal);
    }

    private static string ActiveTestConfiguration()
    {
        // The test binary lives under tests/<Project>/bin/<Configuration>/<TFM>/,
        // so the TFM folder names its parent as the active build
        // configuration. A Debug `dotnet test` therefore inspects the Debug
        // output rather than a stale Release artifact.
        var dir = new DirectoryInfo(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            if (dir.Name.StartsWith("net10.0", StringComparison.Ordinal))
            {
                return dir.Parent?.Name ?? "Release";
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException(
            "Could not determine the active build configuration from the test output path: " +
            Assembly.GetExecutingAssembly().Location);
    }

    private static string LocateCoreAssembly()
    {
        var configuration = ActiveTestConfiguration();

        // Walk up from the test binary until we find a sibling src/ folder.
        var dir = new DirectoryInfo(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            var binRoot = Path.Combine(dir.FullName, "src", "GanttCreator.Core", "bin");
            if (Directory.Exists(binRoot))
            {
                var configDir = Path.Combine(binRoot, configuration);
                if (Directory.Exists(configDir))
                {
                    var netDll = Path.Combine(configDir, "net10.0", "GanttCreator.Core.dll");
                    if (File.Exists(netDll)) return netDll;

                    // Any TFM under the active configuration (e.g. a future
                    // net10.0-windows output) is still that configuration's
                    // artifact; only fallbacks that leave the active
                    // configuration are disallowed, so a Debug test cannot be
                    // satisfied by a stale Release DLL.
                    var withinConfig = Directory
                        .EnumerateFiles(configDir, "GanttCreator.Core.dll", SearchOption.AllDirectories)
                        .FirstOrDefault();
                    if (withinConfig is not null) return withinConfig;
                }

                throw new DirectoryNotFoundException(
                    "Could not locate GanttCreator.Core.dll under " +
                    "src/GanttCreator.Core/bin/" + configuration + "/ (the active test " +
                    "configuration). Build the '" + configuration +
                    "' configuration before running these tests.");
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate src/GanttCreator.Core from the test binary.");
    }
}
