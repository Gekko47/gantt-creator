using System.IO;
using System.Text.Json;
using Xunit;

namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Validates that <c>.vscode/tasks.json</c> conforms to the VS Code tasks
/// 2.0 schema and that the dependency graph is well-formed.
/// </summary>
public sealed class VscodeTasksTests
{
    private static string LocateTasksJson()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(VscodeTasksTests).Assembly.Location)!);
        while (dir is not null)
        {
            var candidate = Path.Combine(dir.FullName, ".vscode", "tasks.json");
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new DirectoryNotFoundException("Could not locate .vscode/tasks.json from the test binary.");
    }

    [Fact]
    public void Tasks_json_uses_VS_Code_object_schema()
    {
        var path = LocateTasksJson();
        using var doc = JsonDocument.Parse(File.ReadAllText(path));

        Assert.True(doc.RootElement.ValueKind == JsonValueKind.Object,
            "tasks.json must be an object { version, tasks }, not a bare array.");

        Assert.True(doc.RootElement.TryGetProperty("version", out var version),
            "tasks.json must have a 'version' property.");
        Assert.Equal("2.0.0", version.GetString());

        Assert.True(doc.RootElement.TryGetProperty("tasks", out var tasks),
            "tasks.json must have a 'tasks' array.");
        Assert.True(tasks.ValueKind == JsonValueKind.Array,
            "'tasks' must be an array.");
    }

    [Fact]
    public void Tasks_json_dependsOn_labels_all_resolve()
    {
        var path = LocateTasksJson();
        using var doc = JsonDocument.Parse(File.ReadAllText(path));

        var tasks = doc.RootElement.GetProperty("tasks");
        var labels = new HashSet<string>(StringComparer.Ordinal);

        foreach (var task in tasks.EnumerateArray())
        {
            if (task.TryGetProperty("label", out var label))
                labels.Add(label.GetString()!);
        }

        foreach (var task in tasks.EnumerateArray())
        {
            if (!task.TryGetProperty("dependsOn", out var dependsOn))
                continue;

            if (dependsOn.ValueKind == JsonValueKind.String)
            {
                var target = dependsOn.GetString()!;
                Assert.True(labels.Contains(target),
                    $"Task '{task.GetProperty("label").GetString()}' dependsOn '{target}' which does not exist.");
            }
            else if (dependsOn.ValueKind == JsonValueKind.Array)
            {
                foreach (var target in dependsOn.EnumerateArray())
                {
                    var targetStr = target.GetString()!;
                    Assert.True(labels.Contains(targetStr),
                        $"Task '{task.GetProperty("label").GetString()}' dependsOn '{targetStr}' which does not exist.");
                }
            }
        }
    }

    [Fact]
    public void Clean_task_has_no_dependencies()
    {
        var path = LocateTasksJson();
        using var doc = JsonDocument.Parse(File.ReadAllText(path));

        var tasks = doc.RootElement.GetProperty("tasks");
        JsonElement? clean = null;

        foreach (var task in tasks.EnumerateArray())
        {
            if (task.TryGetProperty("label", out var label) && label.GetString() == "clean")
            {
                clean = task;
                break;
            }
        }

        Assert.True(clean.HasValue, "A 'clean' task must exist.");
        Assert.False(clean.Value.TryGetProperty("dependsOn", out _),
            "'clean' must not depend on any other task — it must run standalone.");
    }

    [Fact]
    public void Publish_addin_depends_on_build()
    {
        var path = LocateTasksJson();
        using var doc = JsonDocument.Parse(File.ReadAllText(path));

        var tasks = doc.RootElement.GetProperty("tasks");
        JsonElement? publishAddin = null;

        foreach (var task in tasks.EnumerateArray())
        {
            if (task.TryGetProperty("label", out var label) && label.GetString() == "publish-addin")
            {
                publishAddin = task;
                break;
            }
        }

        Assert.True(publishAddin.HasValue, "A 'publish-addin' task must exist.");
        var dependsOn = publishAddin.Value.GetProperty("dependsOn");
        var labels = new List<string>();
        if (dependsOn.ValueKind == JsonValueKind.String)
            labels.Add(dependsOn.GetString()!);
        else if (dependsOn.ValueKind == JsonValueKind.Array)
            labels.AddRange(dependsOn.EnumerateArray().Select(e => e.GetString()!));
        Assert.Contains("build", labels, StringComparer.Ordinal);
    }

    [Fact]
    public void Test_all_depends_on_publish_addin()
    {
        var path = LocateTasksJson();
        using var doc = JsonDocument.Parse(File.ReadAllText(path));

        var tasks = doc.RootElement.GetProperty("tasks");
        JsonElement? testAll = null;

        foreach (var task in tasks.EnumerateArray())
        {
            if (task.TryGetProperty("label", out var label) && label.GetString() == "test-all")
            {
                testAll = task;
                break;
            }
        }

        Assert.True(testAll.HasValue, "A 'test-all' task must exist.");
        var dependsOn = testAll.Value.GetProperty("dependsOn");
        var labels = new List<string>();
        if (dependsOn.ValueKind == JsonValueKind.String)
            labels.Add(dependsOn.GetString()!);
        else if (dependsOn.ValueKind == JsonValueKind.Array)
            labels.AddRange(dependsOn.EnumerateArray().Select(e => e.GetString()!));
        Assert.Contains("publish-addin", labels, StringComparer.Ordinal);
    }
}
