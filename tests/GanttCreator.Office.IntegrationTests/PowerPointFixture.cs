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

/// <summary>What one measurement of the last pasted group reported (R7.4 Phase 0 probe).</summary>
/// <param name="ChildCount">The group's <c>GroupItems.Count</c>.</param>
/// <param name="LeftPt">The group's absolute left edge in points.</param>
/// <param name="TopPt">The group's absolute top edge in points.</param>
/// <param name="WidthPt">The group's width in points.</param>
/// <param name="HeightPt">The group's height in points.</param>
internal sealed record PowerPointGroupBox(int ChildCount, double LeftPt, double TopPt, double WidthPt, double HeightPt);

/// <summary>
/// One text-bearing child's measured box, for the R7.4 text-fidelity evidence.
/// </summary>
/// <param name="Name">The child shape's name.</param>
/// <param name="ShapeWidthPt">The child shape's width in points.</param>
/// <param name="ShapeHeightPt">The child shape's height in points.</param>
/// <param name="HasText">Whether the child carries non-empty text.</param>
/// <param name="TextWidthPt">The measured text bound width in points; zero when there is no text.</param>
/// <param name="TextHeightPt">The measured text bound height in points; zero when there is no text.</param>
/// <param name="WordWrapOn">
/// Whether the text frame wraps (<c>TextFrame2.WordWrap = msoTrue</c>). A wrapped range cannot
/// spill horizontally, so it is the fact that decides whether a bound wider than the shape is a
/// visible overflow or a measurement property (R7.4 audit M-6).
/// </param>
/// <param name="AutoSize">
/// The text frame's <c>TextFrame2.AutoSize</c> value (the <c>MsoAutoSize</c> integer), so a
/// frame that grows to fit its text is distinguishable from a fixed one.
/// </param>
/// <param name="FontSizePt">The range's font size in points; zero when it could not be read.</param>
internal sealed record PowerPointChildTextMetrics(
    string Name,
    double ShapeWidthPt,
    double ShapeHeightPt,
    bool HasText,
    double TextWidthPt,
    double TextHeightPt,
    bool WordWrapOn,
    int AutoSize,
    double FontSizePt);

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
    private readonly List<PowerPoint.Presentation> _extraPresentations = [];
    private readonly List<PowerPoint.ProtectedViewWindow> _protectedWindows = [];
    private readonly PowerPointProxyOwnership _ownership = new();

    private PowerPoint.Application? _application;
    private PowerPoint.Presentation? _presentation;
    private PowerPoint.Slide? _slide;
    private PowerPoint.ShapeRange? _lastPasteRange;
    private PowerPoint.Shape? _lastPastedGroup;
    private PowerPoint.SlideShowWindow? _slideShowWindow;
    private bool _startedByUs;
    private bool _disposed;

    /// <summary>
    /// The PowerPoint build description read from the owned process's file version,
    /// or <c>unknown</c> when no owned process (and therefore no build) exists.
    /// </summary>
    internal string BuildDescription { get; private set; } = "unknown";

    /// <summary>
    /// The process identifiers this fixture started and therefore owns, primary first.
    /// Read by the R7.7 soak to sample per-process counters and to prove the shared
    /// owned-PID manifest recorded each one. Empty when the fixture attached to a
    /// pre-existing host instead of starting one.
    /// </summary>
    internal IReadOnlyList<int> OwnedProcessIds => _ownedProcessIds;

    /// <summary>
    /// Starts or attaches to desktop PowerPoint and creates the test-owned presentation.
    /// </summary>
    /// <param name="withWindow">
    /// <see langword="true"/> for a deck with a document window (the active-slide and view
    /// cases); <see langword="false"/> for a windowless deck (the no-window case).
    /// </param>
    /// <param name="slideCount">The number of blank slides to add; zero exercises the no-slide case.</param>
    /// <exception cref="InvalidOperationException">PowerPoint could not be created.</exception>
    internal void Initialize(bool withWindow = false, int slideCount = 1)
    {
        Dictionary<int, long> before = SnapshotPowerPointProcesses();
        _application = _ownership.Acquire(new PowerPoint.Application());
        AdoptOwnedProcesses(before);
        BuildDescription = DescribeOwnedBuild();

        // ppAlertsNone: ungroup/paste prompts must surface as COM errors, never as a
        // modal dialog that hangs the harness. The alert level is the documented member.
        _application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

        // WithWindow: a windowless deck is the documented no-window case; a windowed deck is
        // what the active-slide, view and slide-show cases exercise.
        PowerPoint.Presentations presentations = _ownership.Acquire(_application.Presentations);
        _presentation = _ownership.Acquire(
            presentations.Add(withWindow ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse));

        PowerPoint.Slides slides = _ownership.Acquire(_presentation.Slides);
        for (int index = 1; index <= slideCount; index++)
        {
            _slide = _ownership.Acquire(slides.Add(index, PowerPoint.PpSlideLayout.ppLayoutBlank));
        }

        _lastPasteRange = null;
        _lastPastedGroup = null;
    }

    /// <summary>
    /// Starts or attaches to desktop PowerPoint WITHOUT creating a presentation, so a live
    /// body can arrange the D-R7.2 zero-deck create arm. The launched process is adopted and
    /// recorded in the shared owned-PID manifest exactly as <see cref="Initialize"/> does, so
    /// the fixture remains the clean teardown owner.
    /// </summary>
    internal void InitializeWithoutPresentation()
    {
        Dictionary<int, long> before = SnapshotPowerPointProcesses();
        _application = _ownership.Acquire(new PowerPoint.Application());
        AdoptOwnedProcesses(before);
        BuildDescription = DescribeOwnedBuild();

        // ppAlertsNone: a quit with a fixture-created deck must not raise a modal prompt.
        _application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;

        _presentation = null;
        _slide = null;
        _lastPasteRange = null;
        _lastPastedGroup = null;
    }

    /// <summary>Adds an in-memory marker text box to the primary slide without saving.</summary>
    /// <param name="text">The marker text.</param>
    /// <exception cref="InvalidOperationException">No primary slide exists.</exception>
    internal void AddUnsavedMarker(string text)
    {
        if (_slide is null)
        {
            throw new InvalidOperationException("Initialize must run before AddUnsavedMarker.");
        }

        PowerPoint.Shapes shapes = _ownership.Acquire(_slide.Shapes);
        PowerPoint.Shape marker = _ownership.Acquire(
            shapes.AddTextbox(
                OfficeCore.MsoTextOrientation.msoTextOrientationHorizontal,
                12,
                12,
                220,
                40));
        PowerPoint.TextFrame frame = _ownership.Acquire(marker.TextFrame);
        PowerPoint.TextRange range = _ownership.Acquire(frame.TextRange);
        range.Text = text;
    }

    /// <summary>Whether the primary deck reports unsaved changes.</summary>
    /// <returns><see langword="true"/> when the host reports the deck as dirty.</returns>
    internal bool PrimaryPresentationHasUnsavedChanges() =>
        _presentation is not null
        && _presentation.Saved == OfficeCore.MsoTriState.msoFalse;

    /// <summary>Whether the primary slide still carries the named text.</summary>
    /// <param name="text">The marker text.</param>
    /// <returns><see langword="true"/> when a shape on the primary slide contains it.</returns>
    internal bool PrimarySlideHasText(string text)
    {
        if (_slide is null)
        {
            return false;
        }

        PowerPoint.Shapes shapes = _ownership.Acquire(_slide.Shapes);
        for (int index = 1; index <= shapes.Count; index++)
        {
            PowerPoint.Shape shape = _ownership.Acquire(shapes[index]);
            if (shape.HasTextFrame != OfficeCore.MsoTriState.msoTrue)
            {
                continue;
            }

            PowerPoint.TextFrame frame = _ownership.Acquire(shape.TextFrame);
            if (frame.HasText != OfficeCore.MsoTriState.msoTrue)
            {
                continue;
            }

            PowerPoint.TextRange range = _ownership.Acquire(frame.TextRange);
            if (string.Equals(range.Text, text, StringComparison.Ordinal))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>The primary slide's shape count; zero when the fixture owns no slide.</summary>
    /// <returns>The live <c>Shapes.Count</c> for the primary slide.</returns>
    internal int PrimarySlideShapeCount()
    {
        if (_slide is null)
        {
            return 0;
        }

        PowerPoint.Shapes shapes = _ownership.Acquire(_slide.Shapes);
        return shapes.Count;
    }

    /// <summary>
    /// Adds a second test-owned presentation, so a live body can prove the target read
    /// resolves the active deck rather than <c>Presentations[1]</c>.
    /// </summary>
    /// <param name="withWindow">Whether the added deck has a document window.</param>
    /// <param name="slideCount">The number of blank slides to add.</param>
    internal void AddPresentation(bool withWindow, int slideCount)
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Initialize must run before AddPresentation.");
        }

        PowerPoint.Presentations presentations = _ownership.Acquire(_application.Presentations);
        PowerPoint.Presentation extra = _ownership.Acquire(
            presentations.Add(withWindow ? OfficeCore.MsoTriState.msoTrue : OfficeCore.MsoTriState.msoFalse));
        _extraPresentations.Add(extra);

        PowerPoint.Slides slides = _ownership.Acquire(extra.Slides);
        for (int index = 1; index <= slideCount; index++)
        {
            _ = _ownership.Acquire(slides.Add(index, PowerPoint.PpSlideLayout.ppLayoutBlank));
        }
    }

    /// <summary>Reads the active window's current <c>PpViewType</c> value.</summary>
    /// <returns>The integer <c>PpViewType</c> the host reports.</returns>
    internal int ActiveViewType()
    {
        PowerPoint.DocumentWindow window = _ownership.Acquire(ActiveWindow());
        PowerPoint.View view = _ownership.Acquire(window.View);
        return (int)view.Type;
    }

    /// <summary>Sets the active window's view and returns what the host reports afterwards.</summary>
    /// <param name="viewType">The <c>PpViewType</c> to request.</param>
    /// <returns>The integer <c>PpViewType</c> the host reports after the request.</returns>
    internal int SetActiveView(PowerPoint.PpViewType viewType)
    {
        PowerPoint.DocumentWindow window = _ownership.Acquire(ActiveWindow());
        window.ViewType = viewType;
        return ActiveViewType();
    }

    /// <summary>Starts a slide show on the test-owned deck.</summary>
    internal void StartSlideShow()
    {
        if (_presentation is null)
        {
            throw new InvalidOperationException("Initialize must run before StartSlideShow.");
        }

        _slideShowWindow = _ownership.Acquire(_presentation.SlideShowSettings.Run());
    }

    /// <summary>Exits a slide show this fixture started, if one is running.</summary>
    internal void ExitSlideShow()
    {
        if (_slideShowWindow is null)
        {
            return;
        }

        try
        {
            _slideShowWindow.View.Exit();
        }
        catch (COMException)
        {
            // Best-effort: the owned process is quit and killed at teardown if the exit fails.
        }
        finally
        {
            _slideShowWindow = null;
        }
    }

    /// <summary>Saves the test-owned primary deck to a temporary file and returns the path.</summary>
    /// <returns>The temporary <c>.pptx</c> path.</returns>
    internal string SavePrimaryPresentationToTemp()
    {
        if (_presentation is null)
        {
            throw new InvalidOperationException("Initialize must run before SavePrimaryPresentationToTemp.");
        }

        string path = Path.Combine(
            Path.GetTempPath(),
            string.Create(CultureInfo.InvariantCulture, $"ganttcreator-r7-2-{Guid.NewGuid():N}.pptx"));
        _presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
        return path;
    }

    /// <summary>Closes the test-owned primary presentation without saving it.</summary>
    internal void ClosePrimaryPresentation()
    {
        if (_presentation is null)
        {
            return;
        }

        try
        {
            _presentation.Saved = OfficeCore.MsoTriState.msoTrue;
            _presentation.Close();
        }
        catch (COMException)
        {
            // Best-effort: the owned process is quit and killed at teardown if the close fails.
        }
        finally
        {
            _presentation = null;
            _slide = null;
        }
    }

    /// <summary>Opens a saved deck read-only in a document window.</summary>
    /// <param name="path">The deck to open.</param>
    internal void OpenReadOnly(string path)
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Initialize must run before OpenReadOnly.");
        }

        PowerPoint.Presentations presentations = _ownership.Acquire(_application.Presentations);
        PowerPoint.Presentation presentation = _ownership.Acquire(
            presentations.Open(
                path,
                ReadOnly: OfficeCore.MsoTriState.msoTrue,
                Untitled: OfficeCore.MsoTriState.msoFalse,
                WithWindow: OfficeCore.MsoTriState.msoTrue));
        _extraPresentations.Add(presentation);
    }

    /// <summary>Opens a saved deck in a protected-view window the fixture owns.</summary>
    /// <param name="path">The deck to open in protected view.</param>
    /// <returns>
    /// Whether the host opened a protected-view window, and the refusal's HResult when it did
    /// not. The reference host refuses this automation route with <c>0x80004005</c> for a
    /// local deck (measured 2026-11-21), so the refusal is a recorded outcome the live body
    /// reports, not a fixture failure.
    /// </returns>
    internal (bool Opened, int? HResult) OpenProtectedView(string path)
    {
        if (_application is null)
        {
            throw new InvalidOperationException("Initialize must run before OpenProtectedView.");
        }

        PowerPoint.ProtectedViewWindows windows = _ownership.Acquire(_application.ProtectedViewWindows);
        try
        {
            _protectedWindows.Add(_ownership.Acquire(windows.Open(path)));
            return (true, null);
        }
        catch (COMException exception)
        {
            return (false, exception.HResult);
        }
    }

    /// <summary>Adopts a host process the production adapter started, so teardown can quit it.</summary>
    /// <param name="before">The <c>POWERPNT</c> snapshot taken before the adapter attached.</param>
    internal void AdoptHostStartedSince(Dictionary<int, long> before)
    {
        _application = _ownership.Acquire(new PowerPoint.Application());
        AdoptOwnedProcesses(before);
        BuildDescription = DescribeOwnedBuild();
        _application.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
    }

    /// <summary>Snapshots the running <c>POWERPNT</c> processes by id and start time.</summary>
    /// <returns>The snapshot keyed by process id.</returns>
    internal static Dictionary<int, long> SnapshotProcesses() => SnapshotPowerPointProcesses();

    private PowerPoint.DocumentWindow ActiveWindow()
    {
        PowerPoint.DocumentWindow? window = _application?.ActiveWindow;
        return window ?? throw new InvalidOperationException("The fixture has no active document window.");
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
    /// Deletes the shape with the given identifier from the test-owned slide, if present.
    /// This is the cleanup route for a group pasted by the PRODUCTION stage: the fixture's
    /// own <see cref="DeletePastedShapes"/> tracks only the paste it performed itself.
    /// </summary>
    /// <param name="shapeId">The PowerPoint shape identifier.</param>
    /// <returns><see langword="true"/> when a shape was deleted.</returns>
    internal bool DeleteShapeById(int shapeId)
    {
        PowerPoint.Shapes shapes;
        try
        {
            shapes = ResolveShapesForLookup();
        }
        catch (InvalidOperationException)
        {
            return false;
        }

        for (int index = 1; index <= shapes.Count; index++)
        {
            PowerPoint.Shape shape = _ownership.Acquire(shapes[index]);
            if (shape.Id != shapeId)
            {
                continue;
            }

            // Deleting the group removes its flattened children too (ADR-0066 D3).
            shape.Delete();
            return true;
        }

        return false;
    }

    /// <summary>
    /// Ungroups the group pasted by the PRODUCTION stage (R7.4 T-3.1 editability check) and
    /// reports the editable child count. The loose shapes become the tracked last paste range,
    /// so <see cref="DeletePastedShapes"/> removes them at the end of the cycle.
    /// </summary>
    /// <param name="shapeId">The PowerPoint shape identifier of the pasted group.</param>
    /// <returns>The number of editable shapes the ungroup produced.</returns>
    internal int UngroupShapeById(int shapeId)
    {
        PowerPoint.Shape group = FindShapeById(shapeId)
            ?? throw new InvalidOperationException(
                string.Create(CultureInfo.InvariantCulture, $"No shape with id {shapeId} on the test slide."));
        PowerPoint.ShapeRange children = _ownership.Acquire(group.Ungroup());
        _lastPasteRange = children;
        _lastPastedGroup = null;
        return children.Count;
    }

    /// <summary>
    /// Measures the text-bearing children of the group pasted by the PRODUCTION stage (R7.4
    /// T-3.1 label-fidelity check), which the fixture's own paste tracking cannot see.
    /// </summary>
    /// <param name="shapeId">The PowerPoint shape identifier of the pasted group.</param>
    /// <returns>One metrics record per group child, in group order.</returns>
    internal IReadOnlyList<PowerPointChildTextMetrics> MeasureGroupTextById(int shapeId)
    {
        PowerPoint.Shape group = FindShapeById(shapeId)
            ?? throw new InvalidOperationException(
                string.Create(CultureInfo.InvariantCulture, $"No shape with id {shapeId} on the test slide."));
        return MeasureGroupText(group);
    }

    /// <summary>Finds the shape with the given identifier on the test-owned or active slide.</summary>
    private PowerPoint.Shape? FindShapeById(int shapeId)
    {
        PowerPoint.Shapes shapes = ResolveShapesForLookup();
        for (int index = 1; index <= shapes.Count; index++)
        {
            PowerPoint.Shape shape = _ownership.Acquire(shapes[index]);
            if (shape.Id == shapeId)
            {
                return shape;
            }
        }

        return null;
    }

    /// <summary>
    /// Resolves the shape collection a production-pasted shape lives in: the fixture's
    /// initial slide, or — in the zero-deck create arm where the transfer created the deck —
    /// the ACTIVE slide the production paste targeted. The active window is read, never
    /// activated.
    /// </summary>
    /// <returns>The resolved shape collection.</returns>
    /// <exception cref="InvalidOperationException">No slide is resolvable.</exception>
    private PowerPoint.Shapes ResolveShapesForLookup()
    {
        PowerPoint.Slide? slide = _slide;
        if (slide is null)
        {
            PowerPoint.DocumentWindow window = _ownership.Acquire(ActiveWindow());
            PowerPoint.View view = _ownership.Acquire(window.View);
            slide = _ownership.Acquire(view.Slide);
        }

        return _ownership.Acquire(slide.Shapes);
    }

    /// <summary>Sets the test-owned presentation's slide size in points (R7.4 probe).</summary>
    /// <param name="widthPt">The slide width in points.</param>
    /// <param name="heightPt">The slide height in points.</param>
    internal void SetSlideSize(double widthPt, double heightPt)
    {
        if (_presentation is null)
        {
            throw new InvalidOperationException("Initialize must run before SetSlideSize.");
        }

        PowerPoint.PageSetup pageSetup = _ownership.Acquire(_presentation.PageSetup);
        pageSetup.SlideWidth = (float)widthPt;
        pageSetup.SlideHeight = (float)heightPt;
    }

    /// <summary>Reads the test-owned presentation's slide size in points (R7.4 probe).</summary>
    /// <returns>The measured slide width and height in points.</returns>
    internal (double WidthPt, double HeightPt) ReadSlideSize()
    {
        if (_presentation is null)
        {
            throw new InvalidOperationException("Initialize must run before ReadSlideSize.");
        }

        PowerPoint.PageSetup pageSetup = _ownership.Acquire(_presentation.PageSetup);
        return (pageSetup.SlideWidth, pageSetup.SlideHeight);
    }

    /// <summary>Measures the last group this fixture pasted (R7.4 probe).</summary>
    /// <returns>The group's child count and absolute box.</returns>
    internal PowerPointGroupBox LastPastedGroupBox()
    {
        PowerPoint.Shape group = _lastPastedGroup
            ?? throw new InvalidOperationException("PasteSpecialShape must run before LastPastedGroupBox.");
        PowerPoint.GroupShapes children = _ownership.Acquire(group.GroupItems);
        return new PowerPointGroupBox(children.Count, group.Left, group.Top, group.Width, group.Height);
    }

    /// <summary>
    /// Route (a) of the R7.4 Step 0 probe: writes the four absolute box members in one pass.
    /// </summary>
    /// <returns>The group box measured after the writes.</returns>
    internal PowerPointGroupBox ApplyRouteAbsoluteBox(double leftPt, double topPt, double widthPt, double heightPt)
    {
        PowerPoint.Shape group = _lastPastedGroup
            ?? throw new InvalidOperationException("PasteSpecialShape must run before ApplyRouteAbsoluteBox.");
        group.Left = (float)leftPt;
        group.Top = (float)topPt;
        group.Width = (float)widthPt;
        group.Height = (float)heightPt;
        return LastPastedGroupBox();
    }

    /// <summary>
    /// Route (b) of the R7.4 Step 0 probe: locks the aspect ratio and sets the width.
    /// </summary>
    /// <returns>The group box measured after the write.</returns>
    internal PowerPointGroupBox ApplyRouteLockAspectWidth(double widthPt)
    {
        PowerPoint.Shape group = _lastPastedGroup
            ?? throw new InvalidOperationException("PasteSpecialShape must run before ApplyRouteLockAspectWidth.");
        group.LockAspectRatio = OfficeCore.MsoTriState.msoTrue;
        group.Width = (float)widthPt;
        return LastPastedGroupBox();
    }

    /// <summary>
    /// Route (c) of the R7.4 Step 0 probe: scales width and height by the same factor.
    /// </summary>
    /// <returns>The group box measured after the scaling.</returns>
    internal PowerPointGroupBox ApplyRouteUniformScale(float factor)
    {
        PowerPoint.Shape group = _lastPastedGroup
            ?? throw new InvalidOperationException("PasteSpecialShape must run before ApplyRouteUniformScale.");
        // Measured 2026-11-21: RelativeToOriginalSize = msoTrue is refused for an msoGroup
        // ("The RelativeToOriginalSize argument applies only to a picture or an OLE object"),
        // so a group scale must be relative to its current size (msoFalse).
        group.ScaleWidth(factor, OfficeCore.MsoTriState.msoFalse, OfficeCore.MsoScaleFrom.msoScaleFromTopLeft);
        group.ScaleHeight(factor, OfficeCore.MsoTriState.msoFalse, OfficeCore.MsoScaleFrom.msoScaleFromTopLeft);
        return LastPastedGroupBox();
    }

    /// <summary>
    /// Measures the text-bearing children of the last pasted group for the R7.4 probe's
    /// label-fidelity evidence. Only the text <em>bound dimensions</em> are recorded; the
    /// probe records them rather than asserting a containment rule whose member semantics
    /// are unprobed.
    /// </summary>
    /// <returns>One metrics record per group child, in group order.</returns>
    internal IReadOnlyList<PowerPointChildTextMetrics> MeasurePastedGroupText()
    {
        PowerPoint.Shape group = _lastPastedGroup
            ?? throw new InvalidOperationException("PasteSpecialShape must run before MeasurePastedGroupText.");
        return MeasureGroupText(group);
    }

    private List<PowerPointChildTextMetrics> MeasureGroupText(PowerPoint.Shape group)
    {
        PowerPoint.GroupShapes children = _ownership.Acquire(group.GroupItems);
        List<PowerPointChildTextMetrics> metrics = [];
        for (int index = 1; index <= children.Count; index++)
        {
            PowerPoint.Shape child = _ownership.Acquire(children[index]);
            double shapeWidthPt = child.Width;
            double shapeHeightPt = child.Height;

            // A child with no text frame is the normal non-text case. The typed getter raises
            // ArgumentException ("The specified value is out of range") for such a child on the
            // reference host (measured 2026-11-21), so the frame is only read when the host says
            // one exists, and the raising-miss encodings are contained besides.
            if (child.HasTextFrame != OfficeCore.MsoTriState.msoTrue)
            {
                metrics.Add(NonText(child.Name, shapeWidthPt, shapeHeightPt));
                continue;
            }

            try
            {
                PowerPoint.TextFrame2 frame = _ownership.Acquire(child.TextFrame2);
                OfficeCore.TextRange2 range = _ownership.Acquire(frame.TextRange);
                string text = range.Text;
                bool hasText = !string.IsNullOrEmpty(text);
                metrics.Add(new PowerPointChildTextMetrics(
                    child.Name,
                    shapeWidthPt,
                    shapeHeightPt,
                    hasText,
                    hasText ? range.BoundWidth : 0,
                    hasText ? range.BoundHeight : 0,
                    frame.WordWrap == OfficeCore.MsoTriState.msoTrue,
                    (int)frame.AutoSize,
                    hasText ? range.Font.Size : 0));
            }
            catch (COMException)
            {
                // A frame whose text range cannot be read is recorded as non-text.
                metrics.Add(NonText(child.Name, shapeWidthPt, shapeHeightPt));
            }
            catch (ArgumentException)
            {
                // Measured raising-miss encoding for a text range that does not exist.
                metrics.Add(NonText(child.Name, shapeWidthPt, shapeHeightPt));
            }
            catch (InvalidOperationException)
            {
                // The third encoding the interop boundary produces for an absent member.
                metrics.Add(NonText(child.Name, shapeWidthPt, shapeHeightPt));
            }
        }

        return metrics;
    }

    private static PowerPointChildTextMetrics NonText(string name, double shapeWidthPt, double shapeHeightPt) =>
        new(name, shapeWidthPt, shapeHeightPt, false, 0, 0, WordWrapOn: false, AutoSize: 0, FontSizePt: 0);

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

        ExitSlideShow();

        foreach (PowerPoint.ProtectedViewWindow protectedWindow in _protectedWindows)
        {
            try
            {
                protectedWindow.Close();
            }
            catch (COMException)
            {
                // Best-effort: the owned process is quit and killed below if the close fails.
            }
        }

        foreach (PowerPoint.Presentation extra in _extraPresentations)
        {
            try
            {
                extra.Saved = OfficeCore.MsoTriState.msoTrue;
                extra.Close();
            }
            catch (COMException)
            {
                // Best-effort: the owned process is quit and killed below if the close fails.
            }
        }

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
        _slideShowWindow = null;
        _extraPresentations.Clear();
        _protectedWindows.Clear();
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
