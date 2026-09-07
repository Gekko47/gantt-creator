using System.Globalization;
using System.Reflection;

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
        Type coreType = typeof(VersionInfo);
        FieldInfo? field = coreType.GetField("SemanticVersion", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        var semantic = (string?)field?.GetValue(null);
        Assert.NotNull(semantic);
        Assert.Matches(@"^\d+\.\d+\.\d+(?:-[0-9A-Za-z.\-]+)?$", semantic);
    }

    [Fact]
    public void InformationalVersion_is_non_empty_and_semver_prefixed()
    {
        // VersionInfo reads AssemblyInformationalVersionAttribute.InformationalVersion
        // from the assembly. Directory.Build.props sets InformationalVersion to
        // "0.0.0+local" in the absence of a real version, so the test
        // environment will see a valid semver value. We assert the prefix is a
        // 3-part semver; we no longer require a hardcoded "0.0.0".
        Type coreType = typeof(VersionInfo);
        FieldInfo? field = coreType.GetField("InformationalVersion", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
        var info = (string?)field?.GetValue(null);
        Assert.NotNull(info);
        Assert.NotEmpty(info);
        Assert.Matches(@"^\d+\.\d+\.\d+(?:[-+]|$)", info!);
    }

    [Theory]
    [InlineData("1.2.3+abc123", "1.2.3")]
    [InlineData("2.0.0-beta.1+dirty", "2.0.0-beta.1")]
    [InlineData("0.0.0-local", "0.0.0")]
    [InlineData("10.5.2", "10.5.2")]
    [InlineData("1.2.3-beta.1", "1.2.3-beta.1")]
    [InlineData("-1.2.3", "0.0.0")]
    [InlineData("+1.2.3", "0.0.0")]
    [InlineData(" 1.2.3", "0.0.0")]
    [InlineData("01.2.3", "0.0.0")]
    [InlineData("1.02.3", "0.0.0")]
    [InlineData("1.2.3-", "0.0.0")]
    [InlineData("1.2", "0.0.0")]
    [InlineData("1.2.3.4", "0.0.0")]
    public void ExtractSemanticVersion_known_formats(string input, string expected)
    {
        Type coreType = typeof(VersionInfo);
        MethodInfo? method = coreType.GetMethod("ExtractSemanticVersion", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        var result = (string?)method?.Invoke(null, [input]);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ExtractSemanticVersion_is_culture_invariant()
    {
        // The numeric parts are validated character-by-character against
        // ASCII digits; no culture-sensitive parse is involved anywhere.
        CultureInfo previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            Type coreType = typeof(VersionInfo);
            MethodInfo? method = coreType.GetMethod("ExtractSemanticVersion", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var result = (string?)method?.Invoke(null, ["1.2.3+abc123"]);
            Assert.Equal("1.2.3", result);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }
}
