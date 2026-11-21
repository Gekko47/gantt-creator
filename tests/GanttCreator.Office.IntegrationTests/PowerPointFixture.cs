using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
/// Test-only PowerPoint host: the typed replacement for the late-bound R6.7 fixture.
/// </summary>
/// <remarks>
/// <para>
/// <b>Typed PIA, no binder.</b> The fixture drives the same members the production adapter
/// does, through <c>Microsoft.Office.Interop.PowerPoint</c>: <c>Presentations.Add</c>,
/// <c>Slides.Add</c> (blank layout), <c>Shapes.PasteSpecial</c> with <c>ppPasteShape</c>,
/// <c>ShapeRange.Count/Width/Height</c>, <c>Shape.Type</c> with <c>msoGroup</c>,
/// <c>ShapeRange.Ungroup</c>, and <c>Application.DisplayAlerts</c>. No <c>dynamic</c> and no
/// <c>Microsoft.CSharp</c> binder call remains (ADR-0069).
/// </para>
/// <para>
/// <b>Ownership mirrors <see cref="OfficeFixture"/> and never destroys a user session.</b>
/// The pre-existing <c>POWERPNT.EXE</c> processes are recorded as user-owned before launch;
/// only the process the launch created is adopted, and it is recorded in the shared owned-PID
/// manifest so <c>verify-office.ps1</c> can sweep it on a timeout. <c>Quit</c> runs only when
/// the fixture started the process; a presentation is always fixture-created and is closed
/// unsaved. Every proxy is held in the shared
/// <see cref="PowerPointProxyOwnership"/> tracker and released once, in reverse order.
/// </para>
/// </remarks>
internal sealed class PowerPointFixture : IAsyncDisposable
{
    private const string _processName = "POWERPNT";

    /// <summary>The PowerPoint <c>msoGroup</c> shape type identifier.</summary>
    private const int _mssoGroupShapeType = 6;

    private static readonly TimeSpan _gracefulExitTimeout = TimeSpan.FromSeconds(10);

    private readonly List<int> _ownedProcessIds = [];
    private readonly Dictionary<int, long> _ownedProcessStartTicks = [];
    private readonly PowerPointProxyOwnership _ownership = new();

    private PowerPoint.Application? _application;
    private PowerPoint.Presentation? _presentation;
    private PowerPoint.Slide? _slide;
    private PowerPoint.ShapeRange? _lastPasteRange;
    private PowerPoint.Shape? _lastPastedGroup;
    private bool _startedByUs;
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
    /// <exception cref="InvalidOperationException">PowerPoint could not be created.</exception>
    internal void Initialize()
    {
        Dictionary<int, long> before = SnapshotPowerPointProcesses();
        _application = _ownership.Acquire(new PowerPoint.Application());
        AdoptOwnedProcesses(before);
        BuildDescription = DescribeOwnedBuild();

        // ppAlertsNone: ungroup/paste prompts must surface as COM errors, never as a
        // modal dialog that hangs the harness. The alert level is the documented member.
        _application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

        // WithWindow = msoFalse: the test-owned presentation has no window, which is the
        // documented way to keep the automation out of the user's view without relying on
        // Application.Visible (PowerPoint does not honour an invisible application).
        PowerPoint.Presentations presentations = _ownership.Acquire(_application.Presentations);
        _presentation = _ownership.Acquire(presentations.Add(OfficeCore.MsoTriState.msoFalse));

        PowerPoint.Slides slides = _ownership.Acquire(_presentation.Slides);
        _slide = _ownership.Acquire(slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank));
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

        PowerPoint.Shapes shapes = _ownership.Acquire(_slide.Shapes);
        PowerPoint.ShapeRange range = _ownership.Acquire(
            shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteShape));
        _lastPasteRange = range;

        int shapeCount = range.Count;
        double widthPt = range.Width;
        double heightPt = range.Height;
        int groupItemCount = 0;
        PowerPoint.Shape? pastedGroup = null;
        List<string> descriptions = [];

        // Measured on the reference host: pasting an Excel group with ppPasteShape returns
        // a range containing the group (msoGroup) AND its flattened children, so the group
        // must be located within the range rather than assumed to be the only item.
        for (int index = 1; index <= shapeCount; index++)
        {
            PowerPoint.Shape item = _ownership.Acquire(range[index]);
            string name = item.Name;
            int shapeType = (int)item.Type;
            double itemWidth = item.Width;
            double itemHeight = item.Height;
            descriptions.Add(string.Create(
                CultureInfo.InvariantCulture,
                $"{index}:name={name}:type={shapeType}:width={itemWidth}:height={itemHeight}"));

            if (shapeType == _mssoGroupShapeType && pastedGroup is null)
            {
                PowerPoint.GroupShapes groupItems = _ownership.Acquire(item.GroupItems);
                groupItemCount = groupItems.Count;
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

        PowerPoint.ShapeRange range = _lastPasteRange;
        if (_lastPastedGroup is not null)
        {
            PowerPoint.ShapeRange children = _ownership.Acquire(_lastPastedGroup.Ungroup());
            int childCount = children.Count;
            _lastPasteRange = children;
            _lastPastedGroup = null;
            return new PowerPointUngroupOutcome(childCount, WasGroup: true);
        }

        return new PowerPointUngroupOutcome(range.Count, WasGroup: false);
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
                _lastPastedGroup.Delete();
            }
            else
            {
                _lastPasteRange.Delete();
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
                // Saved=msoTrue then Close(): the test-owned presentation is discarded, never
                // saved or overwritten, and no save prompt is raised.
                _presentation.Saved = OfficeCore.MsoTriState.msoTrue;
                _presentation.Close();
            }
            catch (COMException)
            {
                // Best-effort: the owned process is polled and killed below if the close fails.
            }
        }

        if (_startedByUs && _application is not null)
        {
            try
            {
                // Only a host this fixture started is quit. A pre-existing PowerPoint and its
                // open presentations are never closed or terminated (work item D2, ADR-0069).
                _application.Quit();
            }
            catch (COMException)
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
        where T : class => _ownership.Acquire(proxy);

    private void AdoptOwnedProcesses(Dictionary<int, long> before)
    {
        Dictionary<int, long> after = SnapshotPowerPointProcesses();
        foreach ((int processId, long startTicks) in after)
        {
            if (!before.ContainsKey(processId))
            {
                _ownedProcessIds.Add(processId);
                _ownedProcessStartTicks[processId] = startTicks;
                OfficeFixture.RecordOwnedProcessId(processId);
            }
        }

        _startedByUs = _ownedProcessIds.Count > 0;
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
        foreach (Process process in Process.GetProcessesByName(_processName))
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
        _ownership.ReleaseAll();
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

        DateTime deadline = DateTime.UtcNow + _gracefulExitTimeout;
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
