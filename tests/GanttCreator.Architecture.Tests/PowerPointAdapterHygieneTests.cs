using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

// artifact-source: verify-quick.ps1 -> 'build Release -warnaserror'
// GanttCreator.Office.dll is inspected from the Release (or Debug) output produced by the
// build step of scripts/verify-quick.ps1 and scripts/verify.ps1, matching PortHygieneTests.
namespace GanttCreator.Architecture.Tests;

/// <summary>
/// R7.1 structural guards for the PowerPoint adapter: no late binding, no activation, no
/// save/close/quit surface and no sleep or <c>Task.Run</c> around COM. Every scan has a
/// positive control, so a green run cannot mean the checker examined nothing.
/// </summary>
/// <remarks>
/// The worksheet activation rule cannot see a <c>slide</c> receiver, which is why the
/// PowerPoint adapter needs its own scan (work item D4). The port-surface rule is reflective
/// rather than textual: a future <c>Save</c> member added to any <c>IPowerPoint*</c> interface
/// fails without a source edit to this test.
/// </remarks>
public sealed class PowerPointAdapterHygieneTests
{
    /// <summary>The R7.1 production files scanned by every source rule.</summary>
    private static readonly string[] _phaseSevenFiles =
    [
        "src/GanttCreator.Office/IPowerPointApplicationPort.cs",
        "src/GanttCreator.Office/PowerPointProxyOwnership.cs",
        "src/GanttCreator.Office/PowerPointProcessSnapshot.cs",
        "src/GanttCreator.Office/PowerPointTransferAdapter.cs",
        // R7.2: the target fact gatherer resolves the active deck through the port and calls
        // the Core decision table; it joins the same no-dynamic/no-activate path.
        "src/GanttCreator.Office/PowerPointTargetResolver.cs",
        // R7.1: the typed replacement for the R6.7 fixture is part of the Phase-7 path,
        // so the no-dynamic and no-sleep rules cover it too.
        "tests/GanttCreator.Office.IntegrationTests/PowerPointFixture.cs",
        // R7.3: the paste port (a primitive-typed declaration) and its stage with the
        // failure classifier. Both join the no-dynamic/no-Microsoft.CSharp/no-activate/
        // no-sleep path; the reflective port-surface scan above covers the new interface.
        "src/GanttCreator.Office/IPowerPointPastePort.cs",
        "src/GanttCreator.Office/PowerPointPasteStage.cs",
        // R7.4: the placement port (a single-group write by identity, no range write) and its
        // stage with the typed refusals. Both join the no-dynamic/no-Microsoft.CSharp/
        // no-activate/no-sleep path; the reflective port-surface scan above covers the new
        // interface automatically because it is an exported IPowerPoint* type.
        "src/GanttCreator.Office/IPowerPointPlacementPort.cs",
        "src/GanttCreator.Office/PowerPointPlacementStage.cs",
        // R7.5: the transfer-failure taxonomy, the measured retry policy, the bounded retry
        // loop, the failure mapper and the paste-residue cleanup. All join the
        // no-dynamic/no-Microsoft.CSharp/no-activate/no-sleep path; the retry loop carries no
        // host call and the cleanup deletes only the add-in's own residue.
        "src/GanttCreator.Office/PowerPointTransferFailure.cs",
        "src/GanttCreator.Office/PowerPointTransientRetryPolicy.cs",
        "src/GanttCreator.Office/PowerPointRetryLoop.cs",
        "src/GanttCreator.Office/PowerPointTransferFailureMapper.cs",
        "src/GanttCreator.Office/PowerPointPasteResidueCleanup.cs",
    ];

    /// <summary>Members the transfer surface must never expose (work item D5).</summary>
    private static readonly string[] _forbiddenPortMembers =
    [
        "Save",
        "SaveAs",
        "Close",
        "Quit",
        "QuitAll",
        "Activate",
    ];

    [Fact]
    public void The_phase_seven_files_exist_and_are_scanned()
    {
        string repositoryRoot = FindRepositoryRoot();

        Assert.All(
            _phaseSevenFiles,
            relative => Assert.True(
                File.Exists(Path.Combine(repositoryRoot, relative)),
                $"The scanned file '{relative}' is missing; a missing file would silently shrink the guard."));
    }

    [Fact]
    public void The_phase_seven_files_contain_no_dynamic() =>
        Assert.Empty(ScanRepositoryFiles(DetectDynamic));

    [Fact]
    public void The_dynamic_detector_flags_a_dynamic_declaration() =>
        Assert.True(DetectDynamic("class Probe { void M() { dynamic value = 1; } }"));

    [Fact]
    public void The_dynamic_detector_accepts_typed_code() =>
        Assert.False(DetectDynamic("class Probe { void M() { object value = 1; } }"));

    [Fact]
    public void The_phase_seven_files_contain_no_microsoft_csharp_reference() =>
        Assert.Empty(ScanRepositoryFiles(DetectMicrosoftCSharp));

    [Fact]
    public void The_microsoft_csharp_detector_flags_a_binder_using() =>
        Assert.True(DetectMicrosoftCSharp("using Microsoft.CSharp.RuntimeBinder; class Probe { }"));

    [Fact]
    public void The_microsoft_csharp_detector_accepts_an_ordinary_using() =>
        Assert.False(DetectMicrosoftCSharp("using System; class Probe { }"));

    [Fact]
    public void The_phase_seven_files_never_activate_a_slide_window_or_presentation() =>
        Assert.Empty(ScanRepositoryFiles(DetectActivation));

    [Fact]
    public void The_activation_detector_flags_an_activate_call() =>
        Assert.True(DetectActivation("class Probe { void M(Probe window) { window.Activate(); } void Activate() { } }"));

    [Fact]
    public void The_activation_detector_accepts_a_read_only_access() =>
        Assert.False(DetectActivation("class Probe { int M(Probe window) { return window.Count; } int Count { get; } }"));

    [Fact]
    public void The_phase_seven_files_contain_no_sleep_or_task_run() =>
        Assert.Empty(ScanRepositoryFiles(DetectSleepOrTaskRun));

    [Fact]
    public void The_sleep_detector_flags_thread_sleep() =>
        Assert.True(DetectSleepOrTaskRun("class Probe { void M() { System.Threading.Thread.Sleep(1); } }"));

    [Fact]
    public void The_sleep_detector_flags_task_run() =>
        Assert.True(DetectSleepOrTaskRun("class Probe { void M() { _ = System.Threading.Tasks.Task.Run(() => 1); } }"));

    [Fact]
    public void The_sleep_detector_accepts_ordinary_calls() =>
        Assert.False(DetectSleepOrTaskRun("class Probe { void M() { _ = string.Concat(\"a\", \"b\"); } }"));

    [Fact]
    public void The_powerpoint_ports_cannot_save_close_quit_or_activate()
    {
        Type[] ports = PowerPointPortInterfaces();
        Assert.NotEmpty(ports);

        List<string> offenders = [];
        foreach (Type port in ports)
        {
            foreach (MemberInfo member in ForbiddenMembers(port))
            {
                offenders.Add(port.Name + "." + member.Name);
            }
        }

        Assert.Empty(offenders);
    }

    [Fact]
    public void The_save_surface_detector_flags_a_save_member()
    {
        MemberInfo offender = Assert.Single(ForbiddenMembers(typeof(IPowerPointBadSurface)));

        Assert.Equal("Save", offender.Name);
    }

    [Fact]
    public void No_powerpoint_port_signature_names_an_interop_type()
    {
        Type[] ports = PowerPointPortInterfaces();
        Assert.NotEmpty(ports);

        List<string> offenders = [];
        foreach (Type port in ports)
        {
            foreach (MethodInfo method in port.GetMethods())
            {
                if (IsInteropTypeName(method.ReturnType.FullName))
                {
                    offenders.Add(port.Name + "." + method.Name + " return type");
                }

                foreach (ParameterInfo parameter in method.GetParameters())
                {
                    if (IsInteropTypeName(parameter.ParameterType.FullName))
                    {
                        offenders.Add(port.Name + "." + method.Name + " parameter " + parameter.Name);
                    }
                }
            }
        }

        Assert.Empty(offenders);
    }

    [Fact]
    public void The_interop_type_detector_flags_an_office_core_type() =>
        Assert.True(IsInteropTypeName("Microsoft.Office.Core.MsoTriState"));

    [Fact]
    public void The_interop_type_detector_accepts_a_primitive() =>
        Assert.False(IsInteropTypeName("System.Int32"));

    private static List<string> ScanRepositoryFiles(Func<string, bool> detector)
    {
        string repositoryRoot = FindRepositoryRoot();
        List<string> offenders = [];

        foreach (string relative in _phaseSevenFiles)
        {
            string path = Path.Combine(repositoryRoot, relative);
            if (!File.Exists(path))
            {
                offenders.Add(relative + " (missing)");
                continue;
            }

            if (detector(File.ReadAllText(path)))
            {
                offenders.Add(relative);
            }
        }

        return offenders;
    }

    private static bool DetectDynamic(string source) =>
        CSharpSyntaxTree.ParseText(source)
            .GetRoot()
            .DescendantTokens()
            .Any(token => token.IsKind(SyntaxKind.IdentifierToken) && token.ValueText == "dynamic");

    private static bool DetectMicrosoftCSharp(string source)
    {
        SyntaxNode root = CSharpSyntaxTree.ParseText(source).GetRoot();

        return root.DescendantNodes()
                   .OfType<UsingDirectiveSyntax>()
                   .Any(directive =>
                       directive.Name?.ToString().StartsWith("Microsoft.CSharp", StringComparison.Ordinal) == true)
            || root.DescendantTokens()
                   .Any(token =>
                       token.IsKind(SyntaxKind.IdentifierToken)
                       && token.ValueText == "RuntimeBinderException");
    }

    private static bool DetectActivation(string source) =>
        CSharpSyntaxTree.ParseText(source)
            .GetRoot()
            .DescendantNodes()
            .OfType<InvocationExpressionSyntax>()
            .Any(invocation =>
                invocation.Expression is MemberAccessExpressionSyntax access
                && access.Name.Identifier.ValueText == "Activate");

    private static bool DetectSleepOrTaskRun(string source) =>
        CSharpSyntaxTree.ParseText(source)
            .GetRoot()
            .DescendantNodes()
            .OfType<InvocationExpressionSyntax>()
            .Any(invocation =>
                invocation.Expression is MemberAccessExpressionSyntax access
                && ((access.Name.Identifier.ValueText == "Sleep"
                        && access.Expression.ToString().EndsWith("Thread", StringComparison.Ordinal))
                    || (access.Name.Identifier.ValueText == "Run"
                        && access.Expression.ToString().EndsWith("Task", StringComparison.Ordinal))));

    private static bool IsInteropTypeName(string? fullName) =>
        fullName is not null
        && (fullName.StartsWith("Microsoft.Office.Interop", StringComparison.Ordinal)
            || fullName.StartsWith("Microsoft.Office.Core", StringComparison.Ordinal));

    private static IEnumerable<MemberInfo> ForbiddenMembers(Type port) =>
        port.GetMembers()
            .Where(member => _forbiddenPortMembers.Contains(member.Name, StringComparer.Ordinal));

    private static Type[] PowerPointPortInterfaces()
    {
        Assembly assembly = OfficeAssembly();
        return
        [
            .. assembly.GetExportedTypes()
                .Where(type => type.IsInterface)
                .Where(type => type.Namespace == "GanttCreator.Office")
                .Where(type => type.Name.StartsWith("IPowerPoint", StringComparison.Ordinal))
                .OrderBy(type => type.FullName, StringComparer.Ordinal),
        ];
    }

    private static Assembly OfficeAssembly()
    {
        string repositoryRoot = FindRepositoryRoot();
        string configurationRoot = Path.Combine(
            repositoryRoot,
            "src",
            "GanttCreator.Office",
            "bin",
            ActiveTestConfiguration());

        Assert.True(
            Directory.Exists(configurationRoot),
            $"The Office build output '{configurationRoot}' does not exist. Build the active configuration first.");

        string? candidate = Directory
            .EnumerateFiles(configurationRoot, "GanttCreator.Office.dll", SearchOption.AllDirectories)
            .Where(path => !IsRefOrPublishArtifact(path))
            .OrderBy(path => path, StringComparer.Ordinal)
            .FirstOrDefault();

        Assert.NotNull(candidate);

        return Assembly.LoadFrom(candidate);
    }

    private static bool IsRefOrPublishArtifact(string path) =>
        path.Contains(
            Path.DirectorySeparatorChar + "ref" + Path.DirectorySeparatorChar,
            StringComparison.Ordinal)
        || path.Contains(
            Path.DirectorySeparatorChar + "publish" + Path.DirectorySeparatorChar,
            StringComparison.Ordinal);

    private static string ActiveTestConfiguration()
    {
        DirectoryInfo? directory = new(
            Path.GetDirectoryName(typeof(PowerPointAdapterHygieneTests).Assembly.Location)!);

        while (directory is not null)
        {
            if (directory.Name.StartsWith("net10.0", StringComparison.Ordinal))
            {
                return directory.Parent?.Name ?? "Release";
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not determine the active build configuration from the test output path.");
    }

    private static string FindRepositoryRoot()
    {
        DirectoryInfo? directory = new(
            Path.GetDirectoryName(typeof(PowerPointAdapterHygieneTests).Assembly.Location)!);

        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "GanttCreator.slnx")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException(
            "Could not locate GanttCreator.slnx above the test binary.");
    }

    /// <summary>A deliberately bad surface the save rule must refuse.</summary>
    private interface IPowerPointBadSurface
    {
        /// <summary>A member the transfer surface must never have.</summary>
        void Save();
    }
}
