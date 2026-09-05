using System;

namespace GanttCreator.Core.Tests;

/// <summary>
/// Tests for <see cref="VersionInfo"/> version string parsing and fallback.
/// </summary>
public sealed class VersionInfoTests
{
    [Fact]
    public void SemanticVersion_extracts_core_from_informational()
    {
        // The version string is computed at compile time via Directory.Build.props.
        // We test the extraction logic directly.
        var coreType = typeof(VersionInfo);
        var field = coreType.GetField("SemanticVersion", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        var semantic = (string?)field?.GetValue(null);
        Assert.NotNull(semantic);
        Assert.Matches(@"^\d+\.\d+\.\d+$", semantic);
    }

    [Fact]
    public void InformationalVersion_is_non_empty_and_contains_semantic_prefix()
    {
        var coreType = typeof(VersionInfo);
        var field = coreType.GetField("InformationalVersion", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        var info = (string?)field?.GetValue(null);
        Assert.NotNull(info);
        Assert.NotEmpty(info!);
        Assert.StartsWith("0.0.0", info); // fallback in test environment
    }

    [Theory]
    [InlineData("1.2.3+abc123", "1.2.3")]
    [InlineData("2.0.0-beta.1+dirty", "2.0.0-beta.1")]
    [InlineData("0.0.0-local", "0.0.0")]
    [InlineData("10.5.2", "10.5.2")]
    public void ExtractSemanticVersion_known_formats(string input, string expected)
    {
        var coreType = typeof(VersionInfo);
        var method = coreType.GetMethod("ExtractSemanticVersion", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var result = (string?)method?.Invoke(null, new object[] { input });
        Assert.Equal(expected, result);
    }
}