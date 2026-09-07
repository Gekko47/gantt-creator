namespace GanttCreator.Core;

/// <summary>
/// Provides the add-in version string, derived from Git at build time
/// with a deterministic fallback for environments without Git.
/// </summary>
public static class VersionInfo
{
    /// <summary>
    /// The full informational version string, suitable for <c>AssemblyInformationalVersion</c>.
    /// Format: <c>semver+commit-hash</c> when Git is available,
    /// otherwise <c>0.0.0-local</c>.
    /// </summary>
    public static readonly string InformationalVersion = ComputeVersion();

    /// <summary>
    /// The short version without build metadata: <c>major.minor.patch</c>,
    /// retaining a prerelease suffix when present (e.g. <c>2.0.0-beta.1</c>).
    /// Unparseable input degrades to <c>0.0.0</c>.
    /// </summary>
    public static readonly string SemanticVersion = ExtractSemanticVersion(InformationalVersion);

    private static string ComputeVersion()
    {
        // Read AssemblyInformationalVersionAttribute.InformationalVersion
        // from this assembly. Directory.Build.props sets InformationalVersion
        // at build time; if the attribute is absent (e.g., a partial-trust
        // host or a stripped assembly), fall back to "0.0.0-local".
        var attr = typeof(VersionInfo).Assembly
            .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
            .FirstOrDefault() as System.Reflection.AssemblyInformationalVersionAttribute;
        return string.IsNullOrEmpty(attr?.InformationalVersion) ? "0.0.0-local" : attr.InformationalVersion;
    }

    private static string ExtractSemanticVersion(string informational)
    {
        if (string.IsNullOrEmpty(informational))
        {
            return "0.0.0";
        }

        // Strip build metadata after '+'
        var plus = informational.IndexOf('+', StringComparison.Ordinal);
        var core = plus >= 0 ? informational[..plus] : informational;

        // Special case: the fallback version "0.0.0-local" is a build-time marker,
        // not a semver prerelease. Treat it as "0.0.0".
        if (core == "0.0.0-local")
        {
            return "0.0.0";
        }

        // Strict semver major.minor.patch[-prerelease]. The numeric parts are
        // validated against ASCII digits only — never int.TryParse, whose
        // default NumberStyles accept signs and whitespace and whose digit
        // handling is culture-sensitive. Leading zeros are invalid semver
        // numeric identifiers and are rejected.
        var dash = core.IndexOf('-', StringComparison.Ordinal);
        var numbers = dash >= 0 ? core[..dash] : core;
        var prerelease = dash >= 0 ? core[(dash + 1)..] : null;

        var parts = numbers.Split('.');
        if (parts.Length != 3
            || !IsStrictSemverNumber(parts[0])
            || !IsStrictSemverNumber(parts[1])
            || !IsStrictSemverNumber(parts[2])
            || (prerelease is not null && prerelease.Length == 0))
        {
            return "0.0.0";
        }

        return prerelease is null ? numbers : $"{numbers}-{prerelease}";
    }

    /// <summary>
    /// A semver numeric identifier: one or more ASCII digits, no sign, no
    /// whitespace, and no leading zeros (except "0" itself).
    /// </summary>
    private static bool IsStrictSemverNumber(string text)
    {
        if (text.Length == 0 || (text.Length > 1 && text[0] == '0'))
        {
            return false;
        }

        foreach (var c in text)
        {
            if (c is < '0' or > '9')
            {
                return false;
            }
        }

        return true;
    }
}
