using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

namespace GanttCreator.Office.IntegrationTests;

/// <summary>
/// What one <c>Shapes.PasteSpecial(ppPasteShape)</c> call produced: the shape count,
/// how many children the pasted range exposes as a group (zero when the host pasted
/// ungrouped shapes), and the pasted bounds.
/// </summary>
/// <param name="ShapeCount">The pasted <c>ShapeRange.Count</c>.</param>
/// <param name="GroupItemCount">The first shape's <c>GroupItems.Count</c>, or zero when it is not a group.</param>
/// <param name="WidthPt">The pasted range width in points.</param>
/// <param name="HeightPt">The pasted range height in points.</param>
/// <param name="ShapeDescriptions">One name/type/size description per pasted shape, in range order.</param>
internal sealed record PowerPointPasteOutcome(
    int ShapeCount,
    int GroupItemCount,
    double WidthPt,
    double HeightPt,
    string[] ShapeDescriptions);

/// <summary>
/// What <see cref="PowerPointFixture.UngroupPastedShapes"/> observed.
/// </summary>
/// <param name="ChildCount">The shapes remaining after the ungroup (or the pasted count when already ungrouped).</param>
/// <param name="WasGroup">Whether the pasted range was a single group that was ungrouped.</param>
internal sealed record PowerPointUngroupOutcome(int ChildCount, bool WasGroup);

/// <summary>
/// Test-only PowerPoint host for the R6.7 copy-after-delete spike.
/// </summary>
/// <remarks>
/// <para>
/// <b>Late-bound COM, deliberately.</b> The work item's D5 default route takes no new
/// package: the fixture resolves <c>PowerPoint.Application</c> through
/// <see cref="Type.GetTypeFromProgID(string)"/> and drives the minimum member surface
/// verified against the Microsoft VBA object-model documentation — <c>Presentations.Add</c>,
/// <c>Slides.Add</c> (blank layout), <c>Shapes.PasteSpecial</c> with <c>ppPasteShape</c>
/// (value 11), <c>ShapeRange.Count/Width/Height</c>, <c>Shape.Type</c> with
/// <c>msoGroup</c> (value 6), <c>ShapeRange.Ungroup</c>, and
/// <c>Application.DisplayAlerts</c> with <c>ppAlertsNone</c> (value 1). It creates no
/// production adapter; R7.1 owns that.
/// </para>
/// <para>
/// <b>Ownership mirrors <see cref="OfficeFixture"/>.</b> The pre-existing
/// <c>POWERPNT.EXE</c> processes are recorded as user-owned before launch; only the
/// process the launch created is polled after <c>Quit()</c> and only it may be killed,
/// after re-confirming its process start time. Every proxy is held in a local and
/// released once, in reverse order. Cleanup is best-effort by design: a release or a
/// close that fails must not escape a <c>finally</c> into the host.
/// </para>
/// </remarks>
internal sealed class PowerPointFixture : IAsyncDisposable
{
    private const string ProcessName = "POWERPNT";

    /// <summary>The PowerPoint <c>msoGroup</c> shape type identifier.</summary>
    private const int MssoGroupShapeType = 6;

    /// <summary>The PowerPoint <c>ppLayoutBlank</c> slide layout identifier.</summary>
    private const int PpLayoutBlank = 12;

    /// <summary>The PowerPoint <c>ppPasteShape</c> clipboard data type identifier.</summary>
    private const int PpPasteShape = 11;

    /// <summary>The PowerPoint <c>ppAlertsNone</c> alert level.</summary>
    private const int PpAlertsNone = 1;

    private static readonly TimeSpan GracefulExitTimeout = TimeSpan.FromSeconds(10);

    private readonly List<int> _ownedProcessIds = [];
    private readonly Dictionary<int, long> _ownedProcessStartTicks = [];
    private readonly List<object> _proxies = [];

    private object? _application;
    private object? _presentation;
    private object? _slide;
    private object? _lastPasteRange;
    private object? _lastPastedGroup;
    private bool _disposed;

    /// <summary>
    /// The PowerPoint build description read from the owned process's file version,
    /// or <c>unknown</c> when no owned process (and therefore no build) exists.
    /// </summary>
    internal string BuildDescription { get; private set; } = "unknown";

    /// <summary>
    /// Starts or attaches to desktop PowerPoint and creates the test-owned
    /// presentation and blank slide the spike pastes into.
    /// </summary>
    /// <exception cref="InvalidOperationException">PowerPoint is not installed, or the presentation could not be created.</exception>
    internal void Initialize()
    {
        Type applicationType = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new InvalidOperationException(
                "PowerPoint is not installed: the PowerPoint.Application ProgID is not registered.");

        Dictionary<int, long> before = SnapshotPowerPointProcesses();
        object instance = Activator.CreateInstance(applicationType)
            ?? throw new InvalidOperationException("Activator.CreateInstance returned null for PowerPoint.Application.");
        _application = Track(instance);
        AdoptOwnedProcesses(before);
        BuildDescription = DescribeOwnedBuild();

        // ppAlertsNone: ungroup/paste prompts must surface as COM errors, never as a
        // modal dialog that hangs the harness. The alert level is the documented member.
        ((dynamic)_application).DisplayAlerts = PpAlertsNone;

        // WithWindow = msoFalse: the test-owned presentation has no window, which is the
        // documented way to keep the automation out of the user's view without relying on
        // Application.Visible (PowerPoint does not honour an invisible application).
        dynamic presentations = Track(((dynamic)_application).Presentations);
        _presentation = Track(presentations.Add(0));

        dynamic slides = Track(((dynamic)_presentation).Slides);
        _slide = Track(slides.Add(1, PpLayoutBlank));
        _lastPasteRange = null;
        _lastPastedGroup = null;
    }

    /// <summary>
    /// Pastes the clipboard as shapes onto the test-owned slide and returns what the
    /// host materialised. Validates the returned range, so a non-throwing COM call is
    /// never mistaken for a successful paste.
    /// </summary>
    /// <returns>The paste outcome.</returns>
    /// <exception cref="InvalidOperationException">The returned range is empty or has non-positive bounds.</exception>
    internal PowerPointPasteOutcome PasteSpecialShape()
    {
        if (_slide is null)
        {
            throw new InvalidOperationException("Initialize must run before PasteSpecialShape.");
        }

        dynamic shapes = Track(((dynamic)_slide).Shapes);
        dynamic range = Track(shapes.PasteSpecial(PpPasteShape));
        _lastPasteRange = range;

        int shapeCount = (int)range.Count;
        double widthPt = (double)range.Width;
        double heightPt = (double)range.Height;
        int groupItemCount = 0;
        object? pastedGroup = null;
        List<string> descriptions = [];

        // Measured on the reference host: pasting an Excel group with ppPasteShape returns
        // a range containing the group (msoGroup) AND its flattened children, so the group
        // must be located within the range rather than assumed to be the only item.
        for (int index = 1; index <= shapeCount; index++)
        {
            dynamic item = Track(range.Item(index));
            string name = (string)item.Name;
            int shapeType = (int)item.Type;
            double itemWidth = (double)item.Width;
            double itemHeight = (double)item.Height;
            descriptions.Add(string.Create(
                CultureInfo.InvariantCulture,
                $"{index}:name={name}:type={shapeType}:width={itemWidth}:height={itemHeight}"));

            if (shapeType == MssoGroupShapeType && pastedGroup is null)
            {
                dynamic groupItems = Track(item.GroupItems);
                groupItemCount = (int)groupItems.Count;
                pastedGroup = item;
            }
        }

        _lastPastedGroup = pastedGroup;
        ValidatePasteOutcome(shapeCount, widthPt, heightPt);
        return new PowerPointPasteOutcome(shapeCount, groupItemCount, widthPt, heightPt, [.. descriptions]);
    }

    /// <summary>
    /// Ungroups the last pasted range when the host pasted a single group, and reports
    /// the resulting editable child count.
    /// </summary>
    /// <returns>The ungroup outcome.</returns>
    internal PowerPointUngroupOutcome UngroupPastedShapes()
    {
        if (_lastPasteRange is null)
        {
            throw new InvalidOperationException("PasteSpecialShape must run before UngroupPastedShapes.");
        }

        dynamic range = _lastPasteRange;
        if (_lastPastedGroup is not null)
        {
            dynamic children = Track(((dynamic)_lastPastedGroup).Ungroup());
            int childCount = (int)children.Count;
            _lastPasteRange = children;
            _lastPastedGroup = null;
            return new PowerPointUngroupOutcome(childCount, WasGroup: true);
        }

        return new PowerPointUngroupOutcome((int)range.Count, WasGroup: false);
    }

    /// <summary>Deletes the last pasted shapes so repeated cycles start clean.</summary>
    internal void DeletePastedShapes()
    {
        if (_lastPasteRange is null)
        {
            return;
        }

        try
        {
            // Deleting the group removes its children too. Deleting the whole range would
            // also try to delete the children that are already covered by the group, which
            // the host refuses.
            if (_lastPastedGroup is not null)
            {
                ((dynamic)_lastPastedGroup).Delete();
            }
            else
            {
                ((dynamic)_lastPasteRange).Delete();
            }
        }
        catch (COMException)
        {
            // Best-effort cleanup: the presentation is closed at teardown regardless, and a
            // delete that fails must not replace the measurement outcome with a cleanup error.
        }
        finally
        {
            _lastPasteRange = null;
            _lastPastedGroup = null;
        }
    }

    /// <summary>
    /// Rejects a paste result that no consumer could use: zero shapes, or bounds that are
    /// not strictly positive. This is the validator the live body leans on so a
    /// non-throwing COM call is never reported as a successful materialisation.
    /// </summary>
    /// <param name="shapeCount">The pasted shape count.</param>
    /// <param name="widthPt">The pasted width in points.</param>
    /// <param name="heightPt">The pasted height in points.</param>
    /// <exception cref="InvalidOperationException">The outcome is unusable.</exception>
    internal static void ValidatePasteOutcome(int shapeCount, double widthPt, double heightPt)
    {
        if (shapeCount <= 0)
        {
            throw new InvalidOperationException(
                string.Create(CultureInfo.InvariantCulture, $"PowerPoint paste produced no shapes (count {shapeCount})."));
        }

        // NaN compares false here on purpose: a host that reports a NaN bound is not usable.
        if (!(widthPt > 0) || !(heightPt > 0))
        {
            throw new InvalidOperationException(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"PowerPoint paste bounds are not positive (width {widthPt}, height {heightPt})."));
        }
    }

    /// <summary>
    /// Whether the process id still names the process this fixture launched, judged by
    /// its recorded start time. A user-owned process that reused a recycled PID is
    /// therefore never mistaken for an owned one.
    /// </summary>
    /// <param name="processId">The process id to check.</param>
    /// <param name="expectedStartTicks">The start time recorded at launch, in UTC ticks.</param>
    /// <returns><see langword="true"/> when the process exists and is the owned one.</returns>
    internal static bool IsOwnedProcessIdentity(int processId, long expectedStartTicks)
    {
        try
        {
            using Process process = Process.GetProcessById(processId);
            return process.StartTime.ToUniversalTime().Ticks == expectedStartTicks;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        catch (Win32Exception)
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public async ValueTask DisposeAsync()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        if (_presentation is not null)
        {
            try
            {
                // Saved=true then Close(): the test-owned presentation is discarded, never
                // saved or overwritten, and no save prompt is raised.
                ((dynamic)_presentation).Saved = true;
                ((dynamic)_presentation).Close();
            }
            catch (COMException)
            {
                // Best-effort: the owned process is polled and killed below if the close fails.
            }
            catch (RuntimeBinderException)
            {
                // Same reason: a member the host refuses must not escape teardown.
            }
        }

        if (_application is not null)
        {
            try
            {
                ((dynamic)_application).Quit();
            }
            catch (COMException)
            {
                // Polled and escalated below.
            }
            catch (RuntimeBinderException)
            {
                // Polled and escalated below.
            }
        }

        ReleaseProxies();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        await EnsureOwnedProcessesExitedAsync().ConfigureAwait(false);
        GC.SuppressFinalize(this);
    }

    private T Track<T>(T proxy)
        where T : notnull
    {
        _proxies.Add(proxy);
        return proxy;
    }

    private void AdoptOwnedProcesses(Dictionary<int, long> before)
    {
        Dictionary<int, long> after = SnapshotPowerPointProcesses();
        foreach ((int processId, long startTicks) in after)
        {
            if (!before.ContainsKey(processId))
            {
                _ownedProcessIds.Add(processId);
                _ownedProcessStartTicks[processId] = startTicks;
            }
        }
    }

    private string DescribeOwnedBuild()
    {
        foreach (int processId in _ownedProcessIds)
        {
            try
            {
                using Process process = Process.GetProcessById(processId);
                FileVersionInfo? version = process.MainModule?.FileVersionInfo;
                if (version is not null)
                {
                    return string.Create(
                        CultureInfo.InvariantCulture,
                        $"PowerPoint {version.FileVersion} ({processId})");
                }
            }
            catch (ArgumentException)
            {
                // The process went away between adoption and the read.
            }
            catch (InvalidOperationException)
            {
                // Same.
            }
            catch (Win32Exception)
            {
                // MainModule is not readable; the build description stays unknown.
            }
        }

        return "unknown";
    }

    private static Dictionary<int, long> SnapshotPowerPointProcesses()
    {
        Dictionary<int, long> processes = [];
        foreach (Process process in Process.GetProcessesByName(ProcessName))
        {
            using (process)
            {
                try
                {
                    processes[process.Id] = process.StartTime.ToUniversalTime().Ticks;
                }
                catch (Win32Exception)
                {
                    // A process that exits or denies access between enumeration and the read
                    // is not adoptable; it is simply absent from the snapshot.
                }
                catch (InvalidOperationException)
                {
                    // Same.
                }
            }
        }

        return processes;
    }

    private void ReleaseProxies()
    {
        for (int index = _proxies.Count - 1; index >= 0; index--)
        {
            object proxy = _proxies[index];
            try
            {
                if (Marshal.IsComObject(proxy))
                {
                    _ = Marshal.ReleaseComObject(proxy);
                }
            }
            catch (ArgumentException)
            {
                // Already released or never a COM object: releasing twice is not a defect
                // worth failing teardown over.
            }
            catch (InvalidComObjectException)
            {
                // Same.
            }
        }

        _proxies.Clear();
        _application = null;
        _presentation = null;
        _slide = null;
        _lastPasteRange = null;
        _lastPastedGroup = null;
    }

    private async Task EnsureOwnedProcessesExitedAsync()
    {
        if (_ownedProcessIds.Count == 0)
        {
            return;
        }

        DateTime deadline = DateTime.UtcNow + GracefulExitTimeout;
        while (DateTime.UtcNow < deadline && AnyOwnedProcessAlive())
        {
            await Task.Delay(TimeSpan.FromMilliseconds(100)).ConfigureAwait(false);
        }

        foreach (int processId in _ownedProcessIds)
        {
            if (!_ownedProcessStartTicks.TryGetValue(processId, out long startTicks)
                || !IsOwnedProcessIdentity(processId, startTicks))
            {
                continue;
            }

            try
            {
                using Process process = Process.GetProcessById(processId);
                process.Kill(entireProcessTree: true);
                await process.WaitForExitAsync().ConfigureAwait(false);
            }
            catch (ArgumentException)
            {
                // Exited between the identity check and the kill.
            }
            catch (InvalidOperationException)
            {
                // Same.
            }
            catch (Win32Exception)
            {
                // The kill was refused; the ownership sweep in verify-office.ps1 reports it.
            }
        }
    }

    private bool AnyOwnedProcessAlive()
    {
        foreach (int processId in _ownedProcessIds)
        {
            if (_ownedProcessStartTicks.TryGetValue(processId, out long startTicks)
                && IsOwnedProcessIdentity(processId, startTicks))
            {
                return true;
            }
        }

        return false;
    }
}
