using System.Globalization;

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
        // We assert the observable public field, not a private method by name.
        var semantic = VersionInfo.SemanticVersion;
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
        var info = VersionInfo.InformationalVersion;
        Assert.NotNull(info);
        Assert.NotEmpty(info);
        Assert.Matches(@"^\d+\.\d+\.\d+(?:[-+]|$)", info);
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
    [InlineData("1.2.3-alpha..1", "0.0.0")]
    [InlineData("1.2.3-.", "0.0.0")]
    [InlineData("1.2.3-01", "0.0.0")]
    [InlineData("1.2.3-beta.01", "0.0.0")]
    [InlineData("1.2.3-alpha!beta", "0.0.0")]
    [InlineData("1.2.3-0", "1.2.3-0")]
    public void ExtractSemanticVersion_known_formats(string input, string expected)
    {
        // Invoke the public normalization contract directly; no reflection on
        // private method names or visibility.
        Assert.Equal(expected, VersionInfo.ExtractSemanticVersion(input));
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
            Assert.Equal("1.2.3", VersionInfo.ExtractSemanticVersion("1.2.3+abc123"));
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }
}
