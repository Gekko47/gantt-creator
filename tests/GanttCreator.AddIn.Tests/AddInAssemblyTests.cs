namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Verifies that the AddIn assembly has the expected types and that
/// it can be loaded from the build output path. This is a smoke test
/// before the first Excel-DNA integration run.
/// </summary>
public class AddInAssemblyTests
{
    private static Assembly LoadAddInAssembly()
    {
        // Walk up from the test binary until we find the AddIn build output.
        var dir = new DirectoryInfo(
            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)!);
        while (dir is not null)
        {
            var addInBin = Path.Combine(dir.FullName, "src", "GanttCreator.AddIn", "bin");
            if (Directory.Exists(addInBin))
            {
                var dll = Path.Combine(addInBin, "Release", "net10.0-windows",
                    "GanttCreator.AddIn.dll");
                if (File.Exists(dll))
                    return Assembly.LoadFrom(dll);
            }
            dir = dir.Parent;
        }
        throw new FileNotFoundException(
            "Could not locate GanttCreator.AddIn.dll in the build output. " +
            "Build the AddIn project before running these tests.");
    }

    [Fact]
    public void AddIn_assembly_loads_successfully()
    {
        var asm = LoadAddInAssembly();
        Assert.NotNull(asm);
        Assert.Equal("GanttCreator.AddIn", asm.GetName().Name);
    }

    [Fact]
    public void AddIn_references_ExcelDna()
    {
        var asm = LoadAddInAssembly();
        var refs = asm.GetReferencedAssemblies();
        Assert.Contains(refs, r => r.Name == "ExcelDna.Integration");
    }

    [Fact]
    public void AddIn_targets_windows_TFM()
    {
        var asm = LoadAddInAssembly();
        var dllPath = asm.Location;

        // The assembly must be in the net10.0-windows output directory.
        Assert.Contains("net10.0-windows", dllPath.Replace('\\', '/'));
    }
}
