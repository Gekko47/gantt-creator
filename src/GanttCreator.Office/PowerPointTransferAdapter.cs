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
