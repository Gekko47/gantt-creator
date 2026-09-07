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
    /// The short semantic version (major.minor.patch) without commit metadata.
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

        // Parse semver: major.minor.patch[-prerelease]
        // Split by '.' but only up to 3 parts to handle dot-separated prerelease
        var parts = core.Split('.', 3);
        if (!(parts.Length == 3 &&
            int.TryParse(parts[0], out _) &&
            int.TryParse(parts[1], out _)))
        {
            return "0.0.0";
        }

        // Parse the third part which is patch[-prerelease]
        var patchPart = parts[2];
        var dash = patchPart.IndexOf('-', StringComparison.Ordinal);
        var patchAndPrerelease = dash >= 0 ? patchPart[..dash] : patchPart;

        // Validate patch is numeric
        if (!int.TryParse(patchAndPrerelease, out _))
        {
            return "0.0.0";
        }

        // Reconstruct with prerelease if present
        var prerelease = dash >= 0 ? patchPart[(dash + 1)..] : null;
        var version = $"{parts[0]}.{parts[1]}.{patchAndPrerelease}";
        if (prerelease != null)
        {
            version += $"-{prerelease}";
        }
        return version;
    }
}
