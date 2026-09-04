using System.Reflection;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Verifies that the AddIn assembly has the expected types and that
/// it can be loaded from the build output path. This is a smoke test
/// before the first Excel-DNA integration run.
/// </summary>
public class AddInAssemblyTests
{
    // The AddIn project is referenced by this test project, so the CLR
    // may already have loaded AddIn.dll from the test output. When that
    // happens, Assembly.LoadFrom returns the already-loaded copy and
    // .Location points at the test output, not the src build output.
    // We therefore resolve the src build directory by walking up from
    // the test binary, independent of any loaded assembly.
    private static string LocateAddInBuildDirectory()
    {
        var dir = new DirectoryInfo(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            var addInBin = Path.Combine(dir.FullName, "src", "GanttCreator.AddIn", "bin");
            var release = Path.Combine(addInBin, "Release", "net10.0-windows");
            if (Directory.Exists(release))
            {
                return release;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException(
            "Could not locate src/GanttCreator.AddIn/bin/Release/net10.0-windows. " +
            "Build the solution before running these tests.");
    }

    private static Assembly LoadAddInAssembly()
    {
        var buildDir = LocateAddInBuildDirectory();
        var dll = Path.Combine(buildDir, "GanttCreator.AddIn.dll");
        if (!File.Exists(dll))
        {
            throw new FileNotFoundException(
                "Could not locate GanttCreator.AddIn.dll in the build output.",
                dll);
        }
        return Assembly.LoadFrom(dll);
    }

    [Fact]
    public void AddIn_assembly_loads_successfully()
    {
        var asm = LoadAddInAssembly();
        Assert.NotNull(asm);
        Assert.Equal("GanttCreator.AddIn", asm.GetName().Name);
    }

    [Fact]
    public void AddIn_packaged_xll_exists()
    {
        // The Excel-DNA packer embeds the managed assemblies (including
        // the ExcelDna.Integration runtime) inside a single packed XLL.
        // This asserts the packaging pipeline actually produced a
        // non-trivial artefact for the x64 target.
        var buildDir = LocateAddInBuildDirectory();
        var xll = Path.Combine(buildDir, "publish", "GanttCreator.AddIn-AddIn64-packed.xll");

        Assert.True(File.Exists(xll), $"Expected packed XLL at '{xll}'.");
        var info = new FileInfo(xll);
        Assert.True(info.Length > 100_000, "Packed XLL is suspiciously small; the four managed assemblies should be embedded.");
    }

    [Fact]
    public void AddIn_targets_windows_TFM()
    {
        // The AddIn build output lives under net10.0-windows; the
        // assembly must resolve from there, not from a plain net10.0
        // output that would indicate the TFM drifted.
        var buildDir = LocateAddInBuildDirectory();
        Assert.Contains("net10.0-windows", buildDir.Replace('\\', '/'), StringComparison.Ordinal);
    }
}
