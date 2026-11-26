using GanttCreator.Office;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The shared substitute for <see cref="IExcelApplicationAdapter"/>.
/// </summary>
/// <remarks>
/// Every member the AddIn calls is here, with a settable fact, an observable counter, and a
/// captured handler a test can raise. The seven hand-written copies that preceded this one each
/// grew their own subset, so one interface change meant seven edits; a file that needs different
/// behaviour sets a property instead of writing another fake.
/// </remarks>
internal sealed class TestExcelApplicationAdapter : IExcelApplicationAdapter
{
    private readonly List<string> _statusBarMessages = [];
    private Action? _workbookStateHandler;
    private Action? _sheetActivatedHandler;
    private Action<string?>? _sheetChangedHandler;

    /// <summary>Gets or sets the answer <see cref="HasActiveWorkbook"/> gives.</summary>
    public bool? ActiveWorkbook { get; set; } = true;

    /// <summary>Gets or sets the answer <see cref="IsStatusBarVisible"/> gives.</summary>
    public bool? StatusBarVisible { get; set; } = true;

    /// <summary>Gets or sets the folder <see cref="ActiveWorkbookFolder"/> reports.</summary>
    public string? WorkbookFolder { get; set; }

    /// <summary>Gets or sets whether <see cref="HasActiveWorkbook"/> throws, as a failed COM probe would.</summary>
    public bool ThrowOnHasActiveWorkbook { get; set; }

    /// <summary>Gets or sets whether <see cref="SubscribeWorkbookStateChanged"/> throws.</summary>
    public bool RefuseWorkbookStateSubscribe { get; set; }

    /// <summary>Gets or sets whether <see cref="SubscribeSheetActivated"/> throws.</summary>
    public bool RefuseSheetActivationSubscribe { get; set; }

    /// <summary>Gets or sets whether disposing the workbook-state subscription throws.</summary>
    public bool ThrowOnWorkbookStateDispose { get; set; }

    /// <summary>Gets how many times the workbook fact was read.</summary>
    public int ActiveWorkbookReads { get; private set; }

    /// <summary>Gets how many workbook-state subscriptions were made.</summary>
    public int WorkbookStateSubscribeCount { get; private set; }

    /// <summary>Gets how many workbook-state subscriptions were disposed.</summary>
    public int WorkbookStateDisposeCount { get; private set; }

    /// <summary>Gets how many sheet-activation subscriptions were made.</summary>
    public int SheetActivationSubscribeCount { get; private set; }

    /// <summary>Gets how many sheet-activation subscriptions were disposed.</summary>
    public int SheetActivationDisposeCount { get; private set; }

    /// <summary>Gets how many times the status bar was cleared.</summary>
    public int ClearStatusBarCount { get; private set; }

    /// <summary>Gets the messages written to the status bar, in order.</summary>
    public IReadOnlyList<string> StatusBarMessages => _statusBarMessages;

    /// <inheritdoc />
    public bool? HasActiveWorkbook()
    {
        ActiveWorkbookReads++;
        return ThrowOnHasActiveWorkbook
            ? throw new InvalidOperationException("the workbook probe failed")
            : ActiveWorkbook;
    }

    /// <inheritdoc />
    public string? ActiveWorkbookFolder() => WorkbookFolder;

    /// <inheritdoc />
    public bool? IsStatusBarVisible() => StatusBarVisible;

    /// <inheritdoc />
    public void ShowStatusBar(string message) => _statusBarMessages.Add(message);

    /// <inheritdoc />
    public void ClearStatusBar() => ClearStatusBarCount++;

    /// <inheritdoc />
    public IDisposable SubscribeSheetChanged(Action<string?> handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        _sheetChangedHandler = handler;
        return new CountingSubscription(() => _sheetChangedHandler = null);
    }

    /// <inheritdoc />
    public IDisposable SubscribeWorkbookStateChanged(Action handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        WorkbookStateSubscribeCount++;
        if (RefuseWorkbookStateSubscribe)
        {
            throw new InvalidOperationException("the host refused the connection");
        }

        _workbookStateHandler = handler;
        return new CountingSubscription(() =>
        {
            _workbookStateHandler = null;
            WorkbookStateDisposeCount++;
            if (ThrowOnWorkbookStateDispose)
            {
                throw new InvalidOperationException("the detach failed");
            }
        });
    }

    /// <inheritdoc />
    public IDisposable SubscribeSheetActivated(Action handler)
    {
        ArgumentNullException.ThrowIfNull(handler);
        SheetActivationSubscribeCount++;
        if (RefuseSheetActivationSubscribe)
        {
            throw new InvalidOperationException("the host refused the connection");
        }

        _sheetActivatedHandler = handler;
        return new CountingSubscription(() =>
        {
            _sheetActivatedHandler = null;
            SheetActivationDisposeCount++;
        });
    }

    /// <summary>Raises one workbook-state change, as Excel would.</summary>
    public void RaiseWorkbookStateChanged() => _workbookStateHandler?.Invoke();

    /// <summary>Raises one sheet activation, as Excel would.</summary>
    public void RaiseSheetActivated() => _sheetActivatedHandler?.Invoke();

    /// <summary>Raises one worksheet cell change, as Excel would.</summary>
    /// <param name="address">The qualified address Excel reports.</param>
    public void RaiseSheetChanged(string? address) => _sheetChangedHandler?.Invoke(address);

    /// <summary>A subscription that reports its first disposal and ignores later ones.</summary>
    private sealed class CountingSubscription(Action onDispose) : IDisposable
    {
        private Action? _onDispose = onDispose;

        public void Dispose()
        {
            Action? onDispose = _onDispose;
            _onDispose = null;
            onDispose?.Invoke();
        }
    }
}
