using System.Reflection;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Verifies that the AddIn assembly has the expected types and that
/// it can be loaded from the build output path. This is a smoke test
/// before the first Excel-DNA integration run.
/// </summary>
public class AddInAssemblyTests
{
    // artifact-source: verify-quick.ps1 -> 'publish AddIn (packed XLL)'
    // is produced by the 'publish AddIn (packed XLL)' step of
    // scripts/verify-quick.ps1 and scripts/verify.ps1 at
    // src/GanttCreator.AddIn/bin/Release/net10.0-windows/publish/ (the
    // .vscode 'test' task depends on 'publish-addin' for the same reason).
    // docs/02-ARCHITECTURE.md build-pipeline artifact contract.
    //
    // The AddIn project is referenced by this test project, so the CLR
    // may already have loaded AddIn.dll from the test output. When that
    // happens, Assembly.LoadFrom returns the already-loaded copy and
    // .Location points at the test output, not the src build output.
    // We therefore resolve the src build directory by walking up from
    // the test binary, independent of any loaded assembly.
    //
    // The active test build configuration is resolved from the test output
    // path (tests/<Project>/bin/<Configuration>/<TFM>/), so a plain
    // `dotnet test` in Debug is not diverted to a stale Release output and
    // vice versa. A candidate directory only counts when the AddIn DLL
    // actually exists in it, so empty or stale directories are skipped
    // instead of failing with a confusing missing-file error downstream.
    // The exception is AddIn_packaged_xll_exists, which asserts the packed
    // XLL produced by 'dotnet publish' in Release only — that test resolves
    // the Release publish path directly via LocateAddInBinDirectory, not
    // through this config walk.
    private static string ActiveTestConfiguration()
    {
        // The test binary lives under tests/<Project>/bin/<Configuration>/<TFM>/,
        // so the TFM folder (the first ancestor whose name starts with the
        // TFM prefix) names its parent as the active build configuration.
        var dir = new DirectoryInfo(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
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

    private static string LocateAddInBinDirectory()
    {
        var dir = new DirectoryInfo(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            var addInBin = Path.Combine(dir.FullName, "src", "GanttCreator.AddIn", "bin");
            if (Directory.Exists(addInBin))
            {
                return addInBin;
            }
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException(
            "Could not locate src/GanttCreator.AddIn/bin. " +
            "Build the solution before running these tests.");
    }

    private static string LocateAddInBuildDirectory()
    {
        var addInBin = LocateAddInBinDirectory();
        var active = ActiveTestConfiguration();

        // The active configuration is searched first and only the active
        // configuration is considered. Fallback entries from Configurations
        // are not added to the search order, and other configurations are
        // not considered. If the active DLL is missing, report the missing
        // active artifact as an error instead of loading another configuration's
        // DLL (which would mask a misconfigured build).
        var candidate = Path.Combine(addInBin, active, "net10.0-windows");
        var dll = Path.Combine(candidate, "GanttCreator.AddIn.dll");
        if (Directory.Exists(candidate) && File.Exists(dll))
        {
            return candidate;
        }

        throw new DirectoryNotFoundException(
            "Could not locate GanttCreator.AddIn.dll under " +
            $"src/GanttCreator.AddIn/bin/{active}/net10.0-windows. " +
            "Build the solution before running these tests.");
    }

    private static Assembly LoadAddInAssembly()
    {
        var buildDir = LocateAddInBuildDirectory();
        var dll = Path.Combine(buildDir, "GanttCreator.AddIn.dll");
        return !File.Exists(dll)
            ? throw new FileNotFoundException(
                "Could not locate GanttCreator.AddIn.dll in the build output.",
                dll)
            : Assembly.LoadFrom(dll);
    }

    [Fact]
    public void AddIn_assembly_loads_successfully()
    {
        Assembly asm = LoadAddInAssembly();
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
        // The packed XLL is produced only by 'dotnet publish' in Release,
        // so resolve the Release publish path directly rather than from the
        // config-resolved buildDir (which may point at Debug and has no
        // publish/ subdir).
        var addInBin = LocateAddInBinDirectory();
        var xll = Path.Combine(
            addInBin, "Release", "net10.0-windows", "publish",
            "GanttCreator.AddIn-AddIn64-packed.xll");

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
