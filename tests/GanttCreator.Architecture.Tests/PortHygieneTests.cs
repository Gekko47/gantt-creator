using System.IO;
using System.Linq;
using System.Reflection;

// artifact-source: verify-quick.ps1 -> 'build Release -warnaserror'
// The GanttCreator.Office.dll consumed here is produced by the
// 'build Release -warnaserror' step of scripts/verify-quick.ps1 and
// scripts/verify.ps1 (docs/02-ARCHITECTURE.md build-pipeline artifact
// contract). Debug output is also accepted so a plain `dotnet test` is not
// artificially red.
namespace GanttCreator.Architecture.Tests;

/// <summary>
/// Enforces two port-hygiene rules that no existing architecture test covered.
/// </summary>
/// <remarks>
/// <para>
/// <b>No duplicate port member.</b> R4.1 D1 requires the two new ports to be
/// subtractive from the fourteen that already exist. A member signature
/// declared on two Office interfaces is a silent ambiguity: a caller's mock
/// satisfies both, and the "which port owns this" question is answered by
/// whichever one happened to be injected. The check reflects over every
/// interface in <c>GanttCreator.Office</c> and fails on a repeated signature.
/// </para>
/// <para>
/// <b>The AddIn names no interop types.</b> R4.1 adds the ports the renderer
/// consumes, and the reason they are primitive is the CS0433 hazard documented
/// on <c>IExcelApplicationAdapter</c>: both <c>ExcelDna.Integration</c> and the
/// Office PIA declare a public <c>Application</c> type, and the AddIn references
/// both. <c>CoreBoundaryTests</c> covers only <c>GanttCreator.Core</c>, so
/// nothing was enforcing the AddIn half of the rule — this closes that gap.
/// </para>
/// <para>
/// Each rule ships with a positive control, so a passing test cannot be an
/// artefact of a checker that never fires.
/// </para>
/// </remarks>
public sealed class PortHygieneTests
{
    private const string OfficeAssemblyFileName = "GanttCreator.Office.dll";

    private static Assembly OfficeAssembly()
    {
        // The architecture test project does not reference GanttCreator.Office,
        // so the assembly is located by path from the repo root exactly as
        // CoreBoundaryTests locates GanttCreator.Core.dll. The active build
        // configuration is taken from the test output path, so a Debug
        // `dotnet test` cannot be satisfied by a stale Release artifact.
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(PortHygieneTests).Assembly.Location)!);
        string configuration = ActiveTestConfiguration();

        while (dir is not null)
        {
            var binRoot = Path.Combine(dir.FullName, "src", "GanttCreator.Office", "bin");
            if (Directory.Exists(binRoot))
            {
                var configDir = Path.Combine(binRoot, configuration);
                if (Directory.Exists(configDir))
                {
                    var candidate = Directory
                        .EnumerateFiles(configDir, OfficeAssemblyFileName, SearchOption.AllDirectories)
                        .Where(path => !IsRefOrPublishArtifact(path))
                        .OrderBy(path => path, StringComparer.Ordinal)
                        .FirstOrDefault();
                    if (candidate is not null)
                    {
                        return Assembly.LoadFrom(candidate);
                    }
                }

                throw new DirectoryNotFoundException(
                    "Could not locate GanttCreator.Office.dll under " +
                    "src/GanttCreator.Office/bin/" + configuration + "/ (the active test " +
                    "configuration). Build the '" + configuration + "' configuration first.");
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate src/GanttCreator.Office from the test binary.");
    }

    /// <summary>
    /// The active test build configuration, named by the TFM folder above the
    /// test binary's output directory.
    /// </summary>
    private static string ActiveTestConfiguration()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(PortHygieneTests).Assembly.Location)!);
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

    private static Type[] OfficePortInterfaces()
    {
        var assembly = OfficeAssembly();
        return
        [
            .. assembly
                .GetExportedTypes()
                // Only this project's own port types. Filtering on the assembly
                // alone is not enough: reflecting the PIA types this assembly
                // references would report Excel's own interfaces (Comments,
                // ListColumn, Shape, ...) as duplicate ports, which says nothing
                // about this codebase.
                .Where(type => type.Namespace is not null
                    && type.Namespace.StartsWith("GanttCreator.", StringComparison.Ordinal))
                .Where(type => type is { IsInterface: true, IsGenericTypeDefinition: false })
                .OrderBy(type => type.FullName, StringComparer.Ordinal)
        ];
    }

    /// <summary>
    /// The signature key two interfaces would collide on: the member name, its
    /// parameter types, and its return type.
    /// </summary>
    /// <remarks>
    /// The return type is part of the key deliberately. Two ports may each
    /// declare a no-argument <c>Repair()</c> as long as they return different
    /// outcome records — that is two distinct contracts a caller resolves by
    /// which port it holds, not an ambiguity. The dangerous case is the one
    /// where name, parameters, <em>and</em> return type all match, because a
    /// single mock then satisfies both interfaces at once.
    /// </remarks>
    private static string SignatureKey(MethodInfo method) =>
        method.Name
        + "(" + string.Join(",", method.GetParameters().Select(p => p.ParameterType.FullName)) + ")"
        + ":" + method.ReturnType.FullName;

    [Fact]
    public void No_two_office_ports_declare_the_same_member_signature()
    {
        // D1 is about *two different* interfaces claiming one signature, which
        // makes a caller's mock satisfy both and leaves "which port owns this"
        // answered by whichever happened to be injected. One interface declaring a
        // signature twice is a compile error, so it is not this rule's concern.
        var claimedBy = new Dictionary<string, string>(StringComparer.Ordinal);
        var duplicates = new List<string>();

        foreach (var type in OfficePortInterfaces())
        {
            var declaredHere = new HashSet<string>(StringComparer.Ordinal);

            foreach (MethodInfo method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            {
                // Property accessors are get_X/set_X pairs; the property itself is
                // the unit of duplication, not its accessors, so they are keyed by
                // the property name instead.
                bool isAccessor = method.IsSpecialName
                    && (method.Name.StartsWith("get_", StringComparison.Ordinal)
                        || method.Name.StartsWith("set_", StringComparison.Ordinal));
                string key = isAccessor
                    ? "property:" + method.Name[4..] + ":" + method.ReturnType.FullName
                    : SignatureKey(method);

                if (!declaredHere.Add(key))
                {
                    continue;
                }

                if (claimedBy.TryGetValue(key, out string? firstOwner))
                {
                    duplicates.Add($"{key} declared by both {firstOwner} and {type.FullName}");
                }
                else
                {
                    claimedBy[key] = type.FullName!;
                }
            }
        }

        Assert.True(
            duplicates.Count == 0,
            "Duplicate port members across GanttCreator.Office interfaces: " +
            string.Join(" | ", duplicates));
    }

    [Fact]
    public void Checker_flags_two_interfaces_declaring_the_same_member()
    {
        // Positive control: two DIFFERENT interfaces declaring one signature is
        // exactly the R4.1 D1 violation, so the checker must flag it and the rule
        // above cannot pass because it never matches anything.
        MethodInfo first = typeof(IPortHygienePositiveControl)
            .GetMethod(nameof(IPortHygienePositiveControl.Reject))!;
        MethodInfo second = typeof(IPortHygieneCollidingControl)
            .GetMethod(nameof(IPortHygieneCollidingControl.Reject))!;

        Assert.NotEqual(first.DeclaringType, second.DeclaringType);
        Assert.Equal(SignatureKey(first), SignatureKey(second));
    }
    [Fact]
    public void AddIn_sources_name_no_office_interop_types()
    {
        var root = FindRepoRoot();
        var addInSource = Path.Combine(root, "src", "GanttCreator.AddIn");
        Assert.True(Directory.Exists(addInSource), "Could not locate src/GanttCreator.AddIn.");

        string[] offenders = [.. Directory
            .EnumerateFiles(addInSource, "*.cs", SearchOption.AllDirectories)
            .Where(path => !HasGeneratedSegment(path))
            .Where(path => ContainsInteropTypeName(File.ReadAllText(path)))
            .Select(path => Path.GetRelativePath(root, path).Replace('\\', '/'))
            .Order(StringComparer.Ordinal)];

        Assert.True(
            offenders.Length == 0,
            "AddIn sources must not name Office interop types (CS0433 discipline): " +
            string.Join(", ", offenders));
    }

    [Fact]
    public void Checker_flags_a_source_naming_an_interop_type()
    {
        // Positive control for the interop scan.
        Assert.True(ContainsInteropTypeName("using Excel = Microsoft.Office.Interop.Excel;"));
        Assert.True(ContainsInteropTypeName("// mentions Microsoft.Office.Interop.PowerPoint.Shape"));
    }

    /// <summary>
    /// Determines whether a source file names an Office interop type. Both the
    /// PIA and the Office core assembly count, since either would drag the
    /// duplicate <c>Application</c> type into the AddIn compilation.
    /// </summary>
    private static bool ContainsInteropTypeName(string source) =>
        source.Contains("Microsoft.Office.Interop", StringComparison.Ordinal)
        || source.Contains("Microsoft.Office.Core", StringComparison.Ordinal);

    private static bool HasGeneratedSegment(string path) =>
        path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Any(segment => segment.Equals("bin", StringComparison.OrdinalIgnoreCase)
                || segment.Equals("obj", StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// Excludes the <c>ref</c> and <c>publish</c> reference assemblies under a
    /// build output, which carry no implementation and would reflect as a second
    /// copy of the same types. The <c>bin</c> segment itself is <em>not</em>
    /// excluded here: this method filters candidates that are already inside a
    /// build output, so treating <c>bin</c> as generated would discard every file
    /// being searched for.
    /// </summary>
    private static bool IsRefOrPublishArtifact(string path) =>
        path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            .Any(segment => segment.Equals("ref", StringComparison.OrdinalIgnoreCase)
                || segment.Equals("publish", StringComparison.OrdinalIgnoreCase));

    private static string FindRepoRoot()
    {
        var dir = new DirectoryInfo(Path.GetDirectoryName(typeof(PortHygieneTests).Assembly.Location)!);
        while (dir is not null)
        {
            if (Directory.Exists(Path.Combine(dir.FullName, "src", "GanttCreator.AddIn")))
            {
                return dir.FullName;
            }

            dir = dir.Parent;
        }

        throw new DirectoryNotFoundException("Could not locate the repository root from the test binary.");
    }
}

/// <summary>
/// A throwaway interface used as half of the duplicate-signature positive
/// control for <see cref="PortHygieneTests"/>. It is declared here rather than in
/// <c>GanttCreator.Office</c> so it can never contribute a real duplicate.
/// </summary>
internal interface IPortHygienePositiveControl
{
    /// <summary>Returns a refusal code for the given reason.</summary>
    /// <param name="reason">The refusal reason.</param>
    /// <returns>The refusal code.</returns>
    int Reject(string reason);
}

/// <summary>
/// The second half of the duplicate-signature positive control: a different
/// interface declaring the identical member signature, which is precisely what
/// the R4.1 D1 rule forbids between two real ports.
/// </summary>
internal interface IPortHygieneCollidingControl
{
    /// <summary>Returns a refusal code for the given reason.</summary>
    /// <param name="reason">The refusal reason.</param>
    /// <returns>The refusal code.</returns>
    int Reject(string reason);
}