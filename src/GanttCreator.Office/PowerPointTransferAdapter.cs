using System.Globalization;
using System.Runtime.InteropServices;
using GanttCreator.Core.Export;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GanttCreator.Office;

/// <summary>
/// The statically typed PowerPoint implementation of <see cref="IPowerPointApplicationPort"/>
/// (R7.1), <see cref="IPowerPointPastePort"/> (R7.3) and <see cref="IPowerPointPlacementPort"/>
/// (R7.4): attach-or-start, active presentation/window/view/slide facts, the typed
/// <c>ppPasteShape</c> paste read, the slide-size fact read, the single-group placement
/// application, and reverse-order release of every proxy it acquired.
/// </summary>
/// <remarks>
/// <para>
/// <b>Typed PIA, no binder.</b> The adapter is statically typed against
/// <c>Microsoft.Office.Interop.PowerPoint</c>; no <c>dynamic</c> member and no
/// <c>Microsoft.CSharp</c> binder call appears in this file (ADR-0069).
/// </para>
/// <para>
/// <b>Ownership is explicit.</b> The application proxy is acquired through
/// <see cref="PowerPointProxyOwnership"/> and released once at <see cref="Dispose"/>.
/// Per-read proxies (collections, presentation, window, view, slide) are acquired through a
/// transient tracker released in a <c>finally</c>, so a repeated read cannot accumulate
/// references.
/// </para>
/// <para>
/// <b>Non-destructive.</b> The adapter never calls <c>Save</c>, <c>SaveAs</c>, <c>Close</c>,
/// <c>Quit</c> or <c>Activate</c>. Starting PowerPoint is announced by the caller; the adapter
/// merely reports <see cref="PowerPointApplicationFacts.StartedByUs"/> so the command can.
/// Releasing the application proxy does not terminate a pre-existing instance.
/// </para>
/// <para>
/// <b>STA only.</b> Every member requires the Office main STA thread. There is no async
/// wrapper and no <c>Task.Run</c> around COM.
/// </para>
/// </remarks>
public sealed class PowerPointTransferAdapter
    : IPowerPointApplicationPort, IPowerPointPastePort, IPowerPointPlacementPort, IDisposable
{
    /// <summary>
    /// The PowerPoint <c>MsoShapeType.msoGroup</c> value (6), measured on the reference host.
    /// The paste result is walked item by item so the group is located by this type, never
    /// assumed to be index 1 (ADR-0066 D3).
    /// </summary>
    private const int _groupShapeTypeIdentifier = 6;

    private readonly PowerPointProxyOwnership _rootOwnership;
    private readonly Action<string>? _technicalRecord;
    private readonly PowerPointTransientRetryPolicy _retryPolicy;

    private PowerPoint.Application? _application;
    private bool _startedByUs;
    private int _hostProcessId;
    private bool _disposed;

    /// <summary>Creates an adapter that has not yet attached to PowerPoint.</summary>
    /// <param name="ownership">
    /// The tracker that owns the application proxy. The caller supplies it so the ownership
    /// discipline is one tested type rather than a private detail of the adapter.
    /// </param>
    /// <param name="technicalRecord">Receives one technical record per refused COM call, or <see langword="null"/>.</param>
    /// <param name="retryPolicy">
    /// The transient filter and attempt cap, or <see langword="null"/> for the documented
    /// candidates. The Step-0 probe supplies the measured set in production.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="ownership"/> is <see langword="null"/>.</exception>
    public PowerPointTransferAdapter(
        PowerPointProxyOwnership ownership,
        Action<string>? technicalRecord = null,
        PowerPointTransientRetryPolicy? retryPolicy = null)
    {
        ArgumentNullException.ThrowIfNull(ownership);
        _rootOwnership = ownership;
        _technicalRecord = technicalRecord;
        _retryPolicy = retryPolicy ?? PowerPointTransientRetryPolicy.DocumentedCandidates;
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The call is idempotent: once attached, a second call reports the current facts without
    /// creating a second application proxy. A refused call reports
    /// <see cref="PowerPointApplicationFacts.Available"/> as <see langword="false"/> with the
    /// HRESULT; it never throws a COM exception at the caller.
    /// </remarks>
    public PowerPointApplicationFacts AttachOrStart()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        if (_application is not null)
        {
            return ReadApplicationFacts();
        }

        int[] before = PowerPointProcessSnapshot.Capture();

        PowerPoint.Application application;
        try
        {
            application = new PowerPoint.Application();
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint attach refused", exception.HResult);
            return new PowerPointApplicationFacts(false, false, 0, 0, null, exception.HResult);
        }

        int[] after = PowerPointProcessSnapshot.Capture();
        _startedByUs = PowerPointProcessSnapshot.DeriveStartedByUs(before, after);
        _hostProcessId = PowerPointProcessSnapshot.ResolveHostProcessId(before, after);
        _application = _rootOwnership.Acquire(application);

        return ReadApplicationFacts();
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The active presentation is <c>Presentations[1]</c>, the active window is read from
    /// <c>Application.ActiveWindow</c>, its view from <c>DocumentWindow.View</c> and its slide
    /// from <c>View.Slide</c>. Nothing is activated. A host with no presentation, no window, no
    /// view slide or a slide show reports the facts; R7.2 decides whether to proceed.
    /// </remarks>
    public PowerPointPresentationFacts ReadPresentationFacts()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        PowerPoint.Application? application = _application;
        if (application is null)
        {
            return new PowerPointPresentationFacts(true, null, false, false, false, 0, null, false, null);
        }

        PowerPointProxyOwnership transient = new(_technicalRecord);
        try
        {
            PowerPoint.Presentations presentations = transient.Acquire(application.Presentations);
            if (presentations.Count == 0)
            {
                return new PowerPointPresentationFacts(true, null, false, false, false, 0, null, false, null);
            }

            PowerPoint.Presentation presentation = transient.Acquire(presentations[1]);
            bool readOnly = presentation.ReadOnly == OfficeCore.MsoTriState.msoTrue;
            PowerPoint.DocumentWindows windows = transient.Acquire(presentation.Windows);
            bool hasWindow = windows.Count > 0;
            int slideCount = transient.Acquire(presentation.Slides).Count;
            bool slideShowActive = transient.Acquire(application.SlideShowWindows).Count > 0;

            int? activeViewType = null;
            int? activeSlideIndex = null;

            if (hasWindow)
            {
                PowerPoint.DocumentWindow? window = application.ActiveWindow;
                if (window is null)
                {
                    hasWindow = false;
                }
                else
                {
                    PowerPoint.View view = transient.Acquire(window.View);
                    activeViewType = (int)view.Type;

                    // View.Slide is declared object by this PIA; the typed pattern match is
                    // the only safe read, and a non-slide view simply reports no index.
                    if (view.Slide is PowerPoint.Slide slide)
                    {
                        activeSlideIndex = transient.Acquire(slide).SlideIndex;
                    }
                }
            }

            return new PowerPointPresentationFacts(
                true,
                null,
                true,
                readOnly,
                hasWindow,
                slideCount,
                activeViewType,
                slideShowActive,
                activeSlideIndex);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint presentation read refused", exception.HResult);
            return new PowerPointPresentationFacts(false, exception.HResult, false, false, false, 0, null, false, null);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The active presentation is resolved through <c>Application.ActiveWindow.Presentation</c>
    /// when a document window exists, so a deck that is not <c>Presentations[1]</c> is not
    /// mistaken for the paste target; a deck with no window falls back to the first
    /// presentation only to report its facts, which the decision table refuses as
    /// <c>NoWindow</c>. <c>ActiveProtectedViewWindow</c> is the on-top protected window, which
    /// is distinct from the protected-window count. Nothing is activated.
    /// </remarks>
    public PowerPointTargetObservation ReadTargetObservation()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        return _application is not { } host
            ? new PowerPointTargetObservation(true, null, 0, 0, false, false, false, null, false, 0, null)
            : PowerPointRetryLoop.Execute(
                _retryPolicy,
                retryAfterMutation: true,
                _ => ReadTargetAttempt(host),
                _technicalRecord,
                "powerpoint target");
    }

    private PowerPointRetryAttempt<PowerPointTargetObservation> ReadTargetAttempt(PowerPoint.Application application)
    {
        PowerPointProxyOwnership transient = new(_technicalRecord);
        try
        {
            PowerPoint.Presentations presentations = ReadStage(
                "powerpoint target presentations",
                () => transient.Acquire(application.Presentations));
            int presentationCount = ReadStage(
                "powerpoint target presentation count",
                () => presentations.Count);
            int protectedViewWindowCount = ReadStage(
                "powerpoint target protected view count",
                () => transient.Acquire(application.ProtectedViewWindows).Count);

            // Measured 2026-11-21 on the reference host: the typed PIA getter for
            // ActiveProtectedViewWindow raises 0x80048240 ("Invalid request") when no
            // protected-view window is open, although the late-bound path returns Nothing.
            // With a zero count no protected window can be on top, so the getter is read only
            // when one exists — which keeps the common state a readable fact rather than a
            // refused observation.
            bool activeProtectedWindowPresent = false;
            if (protectedViewWindowCount > 0)
            {
                PowerPoint.ProtectedViewWindow? activeProtectedWindow = ReadStage(
                    "powerpoint target active protected window",
                    () => application.ActiveProtectedViewWindow);
                activeProtectedWindowPresent = activeProtectedWindow is not null;
                if (activeProtectedWindow is not null)
                {
                    _ = transient.Acquire(activeProtectedWindow);
                }
            }

            bool slideShowActive = ReadStage(
                "powerpoint target slide show count",
                () => transient.Acquire(application.SlideShowWindows).Count > 0);

            if (presentationCount == 0)
            {
                return new PowerPointRetryAttempt<PowerPointTargetObservation>(
                    new PowerPointTargetObservation(
                        true,
                        null,
                        0,
                        protectedViewWindowCount,
                        activeProtectedWindowPresent,
                        false,
                        false,
                        null,
                        slideShowActive,
                        0,
                        null),
                    null,
                    false);
            }

            // Measured 2026-11-21 on the reference host: the typed ActiveWindow getter raises
            // 0x80048240 ("Invalid request") when the host has no active document window (a
            // windowless deck, or a running slide show), although the late-bound path returns
            // Nothing. Absence is the fact the decision table classifies, never a read refusal.
            PowerPoint.DocumentWindow? activeWindow = ReadOptionalStage(
                "powerpoint target active window",
                () => application.ActiveWindow);
            if (activeWindow is not null)
            {
                _ = transient.Acquire(activeWindow);
            }

            PowerPoint.Presentation presentation = ReadStage(
                "powerpoint target presentation",
                () => activeWindow is not null
                    ? transient.Acquire(activeWindow.Presentation)
                    : transient.Acquire(presentations[1]));

            bool readOnly = ReadStage(
                "powerpoint target read only",
                () => presentation.ReadOnly == OfficeCore.MsoTriState.msoTrue);
            int windowCount = ReadStage(
                "powerpoint target windows count",
                () => transient.Acquire(presentation.Windows).Count);
            int slideCount = ReadStage(
                "powerpoint target slides count",
                () => transient.Acquire(presentation.Slides).Count);
            bool hasWindow = activeWindow is not null && windowCount > 0;

            int? activeViewType = null;
            int? activeSlideIndex = null;

            if (activeWindow is not null && windowCount > 0)
            {
                PowerPoint.DocumentWindow window = activeWindow;
                PowerPoint.View view = ReadStage(
                    "powerpoint target view",
                    () => transient.Acquire(window.View));
                activeViewType = ReadStage("powerpoint target view type", () => (int)view.Type);

                // The typed View.Slide getter raises the same 0x80048240 when the deck has no
                // slide; absence is the fact the table classifies as NoSlide/NoActiveSlide.
                PowerPoint.Slide? activeSlide = ReadOptionalStage(
                    "powerpoint target view slide",
                    () => view.Slide as PowerPoint.Slide);
                if (activeSlide is not null)
                {
                    activeSlideIndex = ReadStage(
                        "powerpoint target slide index",
                        () => transient.Acquire(activeSlide).SlideIndex);
                }
            }

            return new PowerPointRetryAttempt<PowerPointTargetObservation>(
                new PowerPointTargetObservation(
                    true,
                    null,
                    presentationCount,
                    protectedViewWindowCount,
                    activeProtectedWindowPresent,
                    readOnly,
                    hasWindow,
                    activeViewType,
                    slideShowActive,
                    slideCount,
                    activeSlideIndex),
                null,
                false);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint target read refused", exception.HResult);
            return new PowerPointRetryAttempt<PowerPointTargetObservation>(
                new PowerPointTargetObservation(false, exception.HResult, 0, 0, false, false, false, null, false, 0, null),
                exception.HResult,
                false);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    /// <inheritdoc/>
    /// <remarks>
    /// <b>The measured form is the group plus its flattened children (ADR-0066 D3).</b> The
    /// returned range is walked item by item so the <c>msoGroup</c> is located by type, never
    /// assumed to be the only item; the expected production form is <c>1 + N</c>. A host
    /// refusal is reported as a refused observation with its HRESULT; an unavailable
    /// application or active slide is reported as a refused observation with no HRESULT, so
    /// the stage can distinguish a missing target from a COM failure (R7.3 D4).
    /// </remarks>
    public PowerPointPasteObservation PasteSpecialShape()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        return _application is not { } host
            ? PowerPointPasteFailureClassifier.RefusedObservation(null)
            : PowerPointRetryLoop.Execute(
                _retryPolicy,
                retryAfterMutation: false,
                _ => PasteAttempt(host),
                _technicalRecord,
                "powerpoint paste");
    }

    private PowerPointRetryAttempt<PowerPointPasteObservation> PasteAttempt(PowerPoint.Application application)
    {
        PowerPointProxyOwnership transient = new(_technicalRecord);
        bool pasteCommitted = false;
        try
        {
            // Measured 2026-11-21: the typed ActiveWindow, View and View.Slide getters raise
            // 0x80048240 ("Invalid request") when the surface is absent, although the
            // late-bound path returns Nothing. Absence is the fact the stage classifies, never
            // a read refusal, so the optional read reports a missing target.
            PowerPoint.DocumentWindow? window = ReadOptionalStage(
                "powerpoint paste active window",
                () => application.ActiveWindow);
            if (window is null)
            {
                return new PowerPointRetryAttempt<PowerPointPasteObservation>(
                    PowerPointPasteFailureClassifier.RefusedObservation(null), null, false);
            }

            _ = transient.Acquire(window);

            PowerPoint.View? view = ReadOptionalStage(
                "powerpoint paste view",
                () => window.View);
            if (view is null)
            {
                return new PowerPointRetryAttempt<PowerPointPasteObservation>(
                    PowerPointPasteFailureClassifier.RefusedObservation(null), null, false);
            }

            _ = transient.Acquire(view);

            PowerPoint.Slide? slide = ReadOptionalStage(
                "powerpoint paste view slide",
                () => view.Slide as PowerPoint.Slide);
            if (slide is null)
            {
                return new PowerPointRetryAttempt<PowerPointPasteObservation>(
                    PowerPointPasteFailureClassifier.RefusedObservation(null), null, false);
            }

            _ = transient.Acquire(slide);

            PowerPoint.Shapes shapes = ReadStage(
                "powerpoint paste shapes",
                () => transient.Acquire(slide.Shapes));
            PowerPoint.ShapeRange range = ReadStage(
                "powerpoint paste special",
                () => transient.Acquire(shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteShape)));

            // The paste has committed. A failure past this point must not re-invoke the paste —
            // that would paste a second group and orphan the first — so the aspect is reported
            // and the caller deletes the residue (PD-3/PD-10).
            pasteCommitted = true;
            return new PowerPointRetryAttempt<PowerPointPasteObservation>(
                ReadPastedRange(transient, range), null, false);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint paste refused", exception.HResult);
            return new PowerPointRetryAttempt<PowerPointPasteObservation>(
                PowerPointPasteFailureClassifier.RefusedObservation(exception.HResult),
                exception.HResult,
                pasteCommitted);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    private PowerPointPasteObservation ReadPastedRange(
        PowerPointProxyOwnership transient,
        PowerPoint.ShapeRange range)
    {
        int count = ReadStage("powerpoint paste range count", () => range.Count);
        double widthPt = ReadStage("powerpoint paste range width", () => range.Width);
        double heightPt = ReadStage("powerpoint paste range height", () => range.Height);

        int groupCount = 0;
        int? groupShapeType = null;
        int groupChildCount = 0;
        string? groupName = null;
        int? groupId = null;

        for (int index = 1; index <= count; index++)
        {
            PowerPoint.Shape item = ReadStage(
                "powerpoint paste range item",
                () => transient.Acquire(range[index]));
            int shapeType = ReadStage(
                "powerpoint paste range item type",
                () => (int)item.Type);

            groupCount += shapeType == _groupShapeTypeIdentifier ? 1 : 0;

            // The first group found is the one whose identity and child count are reported;
            // a second group still increments the count so the validator can refuse it.
            if (groupCount != 1 || groupShapeType is not null)
            {
                continue;
            }

            groupShapeType = shapeType;
            groupChildCount = ReadStage(
                "powerpoint paste group child count",
                () => transient.Acquire(item.GroupItems).Count);
            groupName = ReadStage("powerpoint paste group name", () => item.Name);
            groupId = ReadStage("powerpoint paste group id", () => item.Id);
        }

        return new PowerPointPasteObservation(
            PasteSucceeded: true,
            HResult: null,
            Count: count,
            GroupCount: groupCount,
            GroupShapeType: groupShapeType,
            GroupChildCount: groupChildCount,
            GroupName: groupName,
            GroupId: groupId,
            WidthPt: widthPt,
            HeightPt: heightPt);
    }

    private T ReadStage<T>(string stage, Func<T> read)
    {
        try
        {
            return read();
        }
        catch (COMException exception)
        {
            RecordFailure(stage, exception.HResult);
            throw;
        }
    }

    private T? ReadOptionalStage<T>(string stage, Func<T?> read)
        where T : class
    {
        try
        {
            return read();
        }
        catch (COMException exception)
        {
            // The measured absence code (0x80048240) and every unrecognised code stay absence,
            // as R7.2 established. A measured transient rejection escapes to the operation's
            // retry loop instead, so a busy host is never mistaken for a missing surface (PD-8).
            if (_retryPolicy.IsTransient(exception.HResult))
            {
                throw;
            }

            RecordFailure(stage, exception.HResult);
            return null;
        }
    }

    /// <summary>
    /// Releases the proxies this adapter owns, in reverse acquisition order, exactly once.
    /// </summary>
    /// <remarks>
    /// No <c>Quit</c>, <c>Close</c> or <c>Save</c> is called. Releasing the application proxy
    /// cancels this add-in's reference and leaves a user's PowerPoint running, with its open
    /// presentations untouched.
    /// </remarks>
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _application = null;
        _rootOwnership.ReleaseAll();
        GC.SuppressFinalize(this);
    }

    private PowerPointApplicationFacts ReadApplicationFacts()
    {
        PowerPoint.Application? application = _application;
        if (application is null)
        {
            return new PowerPointApplicationFacts(false, _startedByUs, _hostProcessId, 0, null, null);
        }

        PowerPointProxyOwnership transient = new(_technicalRecord);
        try
        {
            int presentationCount = transient.Acquire(application.Presentations).Count;
            return new PowerPointApplicationFacts(
                true,
                _startedByUs,
                _hostProcessId,
                presentationCount,
                PowerPointProcessSnapshot.DescribeBuild(_hostProcessId),
                null);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint application read refused", exception.HResult);
            return new PowerPointApplicationFacts(false, _startedByUs, _hostProcessId, 0, null, exception.HResult);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The active slide's size lives on its presentation's <c>PageSetup</c>
    /// (<c>SlideWidth</c>/<c>SlideHeight</c>, both <c>Single</c> points). The presentation is
    /// resolved through <c>Application.ActiveWindow.Presentation</c> when a document window
    /// exists, falling back to <c>Presentations[1]</c> only to report a size the decision table
    /// would refuse for lack of a window. Nothing is activated.
    /// </remarks>
    public PowerPointSlideSizeFacts ReadSlideSize()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        return _application is not { } host
            ? new PowerPointSlideSizeFacts(false, null, 0, 0)
            : PowerPointRetryLoop.Execute(
                _retryPolicy,
                retryAfterMutation: true,
                _ => ReadSlideSizeAttempt(host),
                _technicalRecord,
                "powerpoint slide size");
    }

    private PowerPointRetryAttempt<PowerPointSlideSizeFacts> ReadSlideSizeAttempt(PowerPoint.Application application)
    {
        PowerPointProxyOwnership transient = new(_technicalRecord);
        try
        {
            PowerPoint.Presentations presentations = ReadStage(
                "powerpoint slide size presentations",
                () => transient.Acquire(application.Presentations));

            // Measured 2026-11-21: the typed ActiveWindow getter raises 0x80048240 when the host
            // has no document window; absence is a fallback fact, never a read refusal.
            PowerPoint.DocumentWindow? window = ReadOptionalStage(
                "powerpoint slide size active window",
                () => application.ActiveWindow);
            if (window is not null)
            {
                _ = transient.Acquire(window);
            }

            PowerPoint.Presentation presentation = ReadStage(
                "powerpoint slide size presentation",
                () => window is not null
                    ? transient.Acquire(window.Presentation)
                    : transient.Acquire(presentations[1]));

            PowerPoint.PageSetup pageSetup = ReadStage(
                "powerpoint slide size page setup",
                () => transient.Acquire(presentation.PageSetup));
            double widthPt = ReadStage(
                "powerpoint slide size width",
                () => (double)pageSetup.SlideWidth);
            double heightPt = ReadStage(
                "powerpoint slide size height",
                () => (double)pageSetup.SlideHeight);

            return new PowerPointRetryAttempt<PowerPointSlideSizeFacts>(
                new PowerPointSlideSizeFacts(true, null, widthPt, heightPt),
                null,
                false);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint slide size refused", exception.HResult);
            return new PowerPointRetryAttempt<PowerPointSlideSizeFacts>(
                new PowerPointSlideSizeFacts(false, exception.HResult, 0, 0),
                exception.HResult,
                false);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    /// <inheritdoc/>
    /// <remarks>
    /// <b>The provisional absolute resize route (work item Step 0, route (a)).</b> The four box
    /// members — <c>Left</c>, <c>Top</c>, <c>Width</c>, <c>Height</c> — are written from the
    /// placement, so the applied size is the placement's size on both axes and the ratio cannot
    /// drift. The Phase 0 probe confirms or corrects the route before the live gate; if it pins a
    /// relative route instead, only this method changes — the port surface and the stage's refusal
    /// contract are route-independent. Only the located <c>msoGroup</c> is written; the returned
    /// range's flattened children are never touched (D4).
    /// </remarks>
    public PowerPointPlacementObservation ApplyPlacement(int groupId, SlidePlacement placement)
    {
        ArgumentNullException.ThrowIfNull(placement);
        ObjectDisposedException.ThrowIf(_disposed, this);

        PowerPoint.Application? application = _application;
        if (application is null)
        {
            return RefusedPlacement(PowerPointPlacementApplyOutcome.NoSlide, null);
        }

        // Placement writes are idempotent (route (a): four absolute members), so a retry after
        // the mutation re-applies the same box safely (PD-4).
        return PowerPointRetryLoop.Execute(
            _retryPolicy,
            retryAfterMutation: true,
            _ => ApplyPlacementAttempt(application, groupId, placement),
            _technicalRecord,
            "powerpoint placement");
    }

    private PowerPointRetryAttempt<PowerPointPlacementObservation> ApplyPlacementAttempt(
        PowerPoint.Application application,
        int groupId,
        SlidePlacement placement)
    {
        PowerPointProxyOwnership transient = new(_technicalRecord);
        bool mutationCommitted = false;
        try
        {
            PowerPoint.DocumentWindow? window = ReadOptionalStage(
                "powerpoint placement active window",
                () => application.ActiveWindow);
            if (window is null)
            {
                return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                    RefusedPlacement(PowerPointPlacementApplyOutcome.NoSlide, null), null, false);
            }

            _ = transient.Acquire(window);

            PowerPoint.View? view = ReadOptionalStage(
                "powerpoint placement view",
                () => window.View);
            if (view is null)
            {
                return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                    RefusedPlacement(PowerPointPlacementApplyOutcome.NoSlide, null), null, false);
            }

            _ = transient.Acquire(view);

            PowerPoint.Slide? slide = ReadOptionalStage(
                "powerpoint placement view slide",
                () => view.Slide as PowerPoint.Slide);
            if (slide is null)
            {
                return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                    RefusedPlacement(PowerPointPlacementApplyOutcome.NoSlide, null), null, false);
            }

            _ = transient.Acquire(slide);

            PowerPoint.Shapes shapes = ReadStage(
                "powerpoint placement shapes",
                () => transient.Acquire(slide.Shapes));
            PowerPoint.Shape? group = FindGroupById(transient, shapes, groupId);
            if (group is null)
            {
                return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                    RefusedPlacement(PowerPointPlacementApplyOutcome.GroupMissing, null), null, false);
            }

            // Route (a): absolute box writes from the snapped placement. The values are already
            // float-representable, so the Single-typed write is exact. The four writes are the
            // mutation; a transient failure is retried idempotently (PD-4).
            mutationCommitted = true;
            group.Left = (float)placement.Left;
            group.Top = (float)placement.Top;
            group.Width = (float)placement.Width;
            group.Height = (float)placement.Height;

            return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                new PowerPointPlacementObservation(
                    PowerPointPlacementApplyOutcome.Applied,
                    null,
                    group.Left,
                    group.Top,
                    group.Width,
                    group.Height),
                null,
                false);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint placement refused", exception.HResult);
            return new PowerPointRetryAttempt<PowerPointPlacementObservation>(
                RefusedPlacement(PowerPointPlacementApplyOutcome.HostRejected, exception.HResult),
                exception.HResult,
                mutationCommitted);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    private PowerPoint.Shape? FindGroupById(
        PowerPointProxyOwnership transient,
        PowerPoint.Shapes shapes,
        int groupId)
    {
        int count = ReadStage("powerpoint placement shape count", () => shapes.Count);
        for (int index = 1; index <= count; index++)
        {
            PowerPoint.Shape shape = ReadStage(
                "powerpoint placement shape",
                () => transient.Acquire(shapes[index]));
            int id = ReadStage("powerpoint placement shape id", () => shape.Id);
            int shapeType = ReadStage("powerpoint placement shape type", () => (int)shape.Type);
            if (id == groupId && shapeType == _groupShapeTypeIdentifier)
            {
                return shape;
            }
        }

        return null;
    }

    private static PowerPointPlacementObservation RefusedPlacement(
        PowerPointPlacementApplyOutcome outcome,
        int? hresult) =>
        new(outcome, hresult, 0, 0, 0, 0);

    /// <inheritdoc/>
    /// <remarks>
    /// The D-R7.2 creation branch, reached only when no presentation is open and the caller
    /// announces it. The creation is a mutation: a rejected <c>Presentations.Add</c> call never
    /// executed and is retryable, but once the deck exists a post-creation failure is reported
    /// without creating a second deck (the retry loop stops at the boundary).
    /// </remarks>
    public PowerPointPresentationCreation CreatePresentationWithBlankSlide()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        return _application is not { } host
            ? new PowerPointPresentationCreation(false, null, null)
            : PowerPointRetryLoop.Execute(
                _retryPolicy,
                retryAfterMutation: false,
                _ => CreatePresentationAttempt(host),
                _technicalRecord,
                "powerpoint create");
    }

    private PowerPointRetryAttempt<PowerPointPresentationCreation> CreatePresentationAttempt(
        PowerPoint.Application application)
    {
        PowerPointProxyOwnership transient = new(_technicalRecord);
        bool creationCommitted = false;
        try
        {
            PowerPoint.Presentations presentations = ReadStage(
                "powerpoint create presentations",
                () => transient.Acquire(application.Presentations));
            PowerPoint.Presentation presentation = ReadStage(
                "powerpoint create presentation",
                () => transient.Acquire(presentations.Add(OfficeCore.MsoTriState.msoTrue)));
            creationCommitted = true;
            PowerPoint.Slides slides = ReadStage(
                "powerpoint create slides",
                () => transient.Acquire(presentation.Slides));
            PowerPoint.Slide slide = ReadStage(
                "powerpoint create slide",
                () => transient.Acquire(slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank)));

            return new PowerPointRetryAttempt<PowerPointPresentationCreation>(
                new PowerPointPresentationCreation(true, slide.SlideIndex, null),
                null,
                false);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint create refused", exception.HResult);
            return new PowerPointRetryAttempt<PowerPointPresentationCreation>(
                new PowerPointPresentationCreation(false, null, exception.HResult),
                exception.HResult,
                creationCommitted);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    /// <inheritdoc/>
    /// <remarks>
    /// The R7.5 paste-residue cleanup (PD-10): delete the located <c>msoGroup</c> by id, never
    /// the returned range that also lists its flattened children (ADR-0066 D3). A cleanup delete
    /// is attempted once and its refusal reported — it is never retried.
    /// </remarks>
    public PowerPointGroupDeleteResult DeletePastedGroup(int groupId)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        PowerPoint.Application? application = _application;
        if (application is null)
        {
            return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.GroupMissing, null);
        }

        PowerPointProxyOwnership transient = new(_technicalRecord);
        try
        {
            PowerPoint.DocumentWindow? window = ReadOptionalStage(
                "powerpoint residue delete active window",
                () => application.ActiveWindow);
            if (window is null)
            {
                return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.GroupMissing, null);
            }

            _ = transient.Acquire(window);

            PowerPoint.View? view = ReadOptionalStage(
                "powerpoint residue delete view",
                () => window.View);
            if (view is null)
            {
                return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.GroupMissing, null);
            }

            _ = transient.Acquire(view);

            PowerPoint.Slide? slide = ReadOptionalStage(
                "powerpoint residue delete view slide",
                () => view.Slide as PowerPoint.Slide);
            if (slide is null)
            {
                return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.GroupMissing, null);
            }

            _ = transient.Acquire(slide);

            PowerPoint.Shapes shapes = ReadStage(
                "powerpoint residue delete shapes",
                () => transient.Acquire(slide.Shapes));
            PowerPoint.Shape? group = FindGroupById(transient, shapes, groupId);
            if (group is null)
            {
                return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.GroupMissing, null);
            }

            // Deleting the group removes its flattened children too (ADR-0066 D3).
            group.Delete();
            return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.Deleted, null);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint residue delete refused", exception.HResult);
            return new PowerPointGroupDeleteResult(PowerPointGroupDeleteOutcome.HostRejected, exception.HResult);
        }
        finally
        {
            transient.ReleaseAll();
        }
    }

    private void RecordFailure(string stage, int hresult)
    {
        _technicalRecord?.Invoke(
            stage + " hresult=0x" + hresult.ToString("X8", CultureInfo.InvariantCulture));
    }
}
