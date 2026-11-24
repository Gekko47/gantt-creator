namespace GanttCreator.Raster.Tests;

/// <summary>A unique temporary directory that is removed on dispose.</summary>
internal sealed class TestTempDirectory : IDisposable
{
    /// <summary>Creates a unique temporary directory with a random name.</summary>
    public TestTempDirectory()
    {
        Path = System.IO.Path.Combine(
            System.IO.Path.GetTempPath(),
            "gantt-raster-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path);
    }

    /// <summary>Gets the directory path.</summary>
    public string Path { get; }

    /// <summary>Combines a file name with the directory path.</summary>
    /// <param name="name">The file name.</param>
    /// <returns>The absolute file path.</returns>
    public string File(string name) => System.IO.Path.Combine(Path, name);

    /// <inheritdoc />
    public void Dispose()
    {
        if (!Directory.Exists(Path))
        {
            return;
        }

        foreach (string file in Directory.GetFiles(Path, "*", SearchOption.AllDirectories))
        {
            System.IO.File.SetAttributes(file, FileAttributes.Normal);
        }

        Directory.Delete(Path, true);
        GC.SuppressFinalize(this);
    }
}
