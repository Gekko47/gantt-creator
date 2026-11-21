using System.Globalization;
using System.Runtime.InteropServices;
using OfficeCore = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace GanttCreator.Office;

/// <summary>
/// The statically typed PowerPoint implementation of <see cref="IPowerPointApplicationPort"/>
/// (R7.1): attach-or-start, active presentation/window/view/slide facts, and reverse-order
/// release of every proxy it acquired.
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
public sealed class PowerPointTransferAdapter : IPowerPointApplicationPort, IDisposable
{
    private readonly PowerPointProxyOwnership _rootOwnership;
    private readonly Action<string>? _technicalRecord;

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
    /// <exception cref="ArgumentNullException"><paramref name="ownership"/> is <see langword="null"/>.</exception>
    public PowerPointTransferAdapter(
        PowerPointProxyOwnership ownership,
        Action<string>? technicalRecord = null)
    {
        ArgumentNullException.ThrowIfNull(ownership);
        _rootOwnership = ownership;
        _technicalRecord = technicalRecord;
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

        PowerPoint.Application? application = _application;
        if (application is null)
        {
            return new PowerPointTargetObservation(true, null, 0, 0, false, false, false, null, false, 0, null);
        }

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
                return new PowerPointTargetObservation(
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
                    null);
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

            return new PowerPointTargetObservation(
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
                activeSlideIndex);
        }
        catch (COMException exception)
        {
            RecordFailure("powerpoint target read refused", exception.HResult);
            return new PowerPointTargetObservation(false, exception.HResult, 0, 0, false, false, false, null, false, 0, null);
        }
        finally
        {
            transient.ReleaseAll();
        }
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

    private void RecordFailure(string stage, int hresult)
    {
        _technicalRecord?.Invoke(
            stage + " hresult=0x" + hresult.ToString("X8", CultureInfo.InvariantCulture));
    }
}
