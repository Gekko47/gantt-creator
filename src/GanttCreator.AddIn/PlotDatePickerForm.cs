using System.Diagnostics.CodeAnalysis;

namespace GanttCreator.AddIn;

/// <summary>
/// The values the plot-date picker needs to open (R5.1a): the month to show, the
/// bounds the calendar may offer, and the caption naming the end being edited.
/// </summary>
/// <param name="Title">The window title, naming the end being edited.</param>
/// <param name="Initial">The date the month view opens on and pre-selects.</param>
/// <param name="Minimum">The earliest selectable date, or null for no lower bound.</param>
/// <param name="Maximum">The latest selectable date, or null for no upper bound.</param>
internal sealed record PlotDatePickerRequest(
    string Title,
    DateOnly Initial,
    DateOnly? Minimum,
    DateOnly? Maximum);

/// <summary>
/// The modal popup behind the two plot-date calendar buttons (R5.1a).
/// </summary>
/// <remarks>
/// <para>
/// One native <see cref="MonthCalendar"/> in a small tool-window frame. Month and
/// year scrolling, wheel and keyboard navigation, screen-reader support and
/// high-contrast rendering come from the platform; a custom-drawn grid would have
/// to re-earn all of them.
/// </para>
/// <para>
/// The frame is a normal tool window rather than a borderless popup because a
/// borderless <b>modal</b> form has no close affordance: it swallows outside
/// clicks, so it would sit on screen until a key was pressed. A tool window keeps
/// the close box, <c>Escape</c> and the caption that names the end being edited.
/// </para>
/// <para>
/// Runs on the Excel main STA thread via <c>ShowDialog</c>, fully offline. The
/// result is a <see cref="DateOnly"/> so no time component can leak into storage.
/// </para>
/// </remarks>
internal sealed class PlotDatePickerForm : Form
{
    private readonly MonthCalendar _calendar = new()
    {
        MaxSelectionCount = 1,
        ShowToday = true,
        ShowTodayCircle = true,
        ShowWeekNumbers = false,
        TabIndex = 0,
    };

    /// <summary>
    /// Initializes the popup.
    /// </summary>
    /// <param name="title">The window title, naming the end being edited.</param>
    /// <param name="initial">The date the month view opens on and pre-selects.</param>
    /// <param name="min">The earliest selectable date, or null for no lower bound.</param>
    /// <param name="max">The latest selectable date, or null for no upper bound.</param>
    /// <exception cref="ArgumentException"><paramref name="title"/> is null or whitespace.</exception>
    internal PlotDatePickerForm(string title, DateOnly initial, DateOnly? min, DateOnly? max)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(title);
        InitializeComponent(title, initial, min, max);
    }

    /// <summary>Gets the picked date when the user confirmed, or null when cancelled.</summary>
    internal DateOnly? SelectedDate { get; private set; }

    /// <summary>
    /// Shows the picker for one request and returns the picked date.
    /// </summary>
    /// <param name="request">The month, bounds and caption to open with.</param>
    /// <returns>The picked date, or null when the user cancelled.</returns>
    /// <remarks>
    /// The production entry point behind <c>RibbonStateService.ShowPlotDatePicker</c>.
    /// Keeping the form construction here means the service tests substitute one
    /// seam that returns a date, and never build a WinForms control off the STA
    /// thread.
    /// </remarks>
    internal static DateOnly? Show(PlotDatePickerRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        using PlotDatePickerForm dialog = new(
            request.Title, request.Initial, request.Minimum, request.Maximum);
        return dialog.ShowDialog() == DialogResult.OK ? dialog.SelectedDate : null;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _calendar.Dispose();
        }

        base.Dispose(disposing);
    }

    [SuppressMessage(
        "Microsoft.Globalization",
        "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void InitializeComponent(string title, DateOnly initial, DateOnly? min, DateOnly? max)
    {
        // Each bound is applied on its own: the caller supplies at most one, so a
        // MinDate greater than MaxDate - which the control throws on - cannot be
        // constructed here.
        DateTime initialDate = initial.ToDateTime(TimeOnly.MinValue);
        if (min is DateOnly low)
        {
            _calendar.MinDate = low.ToDateTime(TimeOnly.MinValue);
        }

        if (max is DateOnly high)
        {
            _calendar.MaxDate = high.ToDateTime(TimeOnly.MinValue);
        }

        if (initialDate < _calendar.MinDate)
        {
            initialDate = _calendar.MinDate;
        }

        if (initialDate > _calendar.MaxDate)
        {
            initialDate = _calendar.MaxDate;
        }

        _calendar.SetDate(initialDate);
        _calendar.AccessibleName = title;
        _calendar.AccessibleDescription =
            "Choose a date with the arrow keys; Page Up and Page Down change month.";
        _calendar.DateSelected += OnDateSelected;
        _calendar.Location = new Point(0, 0);

        Text = title;
        FormBorderStyle = FormBorderStyle.FixedToolWindow;
        ShowInTaskbar = false;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowIcon = false;
        StartPosition = FormStartPosition.Manual;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Font = SystemFonts.MessageBoxFont;
        KeyPreview = true;
        Controls.Add(_calendar);

        ApplyTheme();
        Shown += (_, _) => RepositionAtCursor();
        KeyDown += OnKeyDown;
    }

    /// <summary>
    /// Tints the calendar to match the Office palette, or leaves the system
    /// colours alone under Windows high contrast.
    /// </summary>
    /// <remarks>
    /// High contrast is an accessibility contract, not a preference: overriding
    /// the title bar colour there can render the calendar illegible for exactly
    /// the users who need the contrast. The palette comes from the same resolver
    /// the ribbon glyphs use, so the popup and the button that opened it agree.
    /// </remarks>
    private void ApplyTheme()
    {
        if (SystemInformation.HighContrast)
        {
            return;
        }

        bool dark = RibbonGlyphEnvironment.Current == RibbonGlyphPalette.Dark;
        _calendar.TitleBackColor = dark ? Color.FromArgb(0x2B, 0x2B, 0x2B) : Color.FromArgb(0xF3, 0xF3, 0xF3);
        _calendar.TitleForeColor = dark ? Color.FromArgb(0xF2, 0xF2, 0xF2) : Color.FromArgb(0x1F, 0x1F, 0x1F);
        _calendar.BackColor = dark ? Color.FromArgb(0x1F, 0x1F, 0x1F) : Color.White;
        _calendar.TrailingForeColor = dark ? Color.FromArgb(0x9A, 0x9A, 0x9A) : Color.FromArgb(0x80, 0x80, 0x80);
    }

    /// <summary>
    /// Drops the popup below the ribbon button that opened it, clamped inside the
    /// working area of the screen the cursor is on.
    /// </summary>
    private void RepositionAtCursor()
    {
        Point cursor = Cursor.Position;
        Rectangle area = Screen.FromPoint(cursor).WorkingArea;
        int x = Math.Clamp(cursor.X, area.Left, Math.Max(area.Left, area.Right - Width));
        int y = Math.Clamp(cursor.Y, area.Top, Math.Max(area.Top, area.Bottom - Height));
        Location = new Point(x, y);
    }

    private void OnDateSelected(object? sender, DateRangeEventArgs e)
    {
        SelectedDate = DateOnly.FromDateTime(e.Start);
        DialogResult = DialogResult.OK;
        Close();
    }

    private void OnKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
        {
            SelectedDate = null;
            DialogResult = DialogResult.Cancel;
            Close();
            e.Handled = true;
            return;
        }

        if (e.KeyCode == Keys.Enter)
        {
            // Enter confirms the highlighted day when the calendar grid has focus.
            SelectedDate = DateOnly.FromDateTime(_calendar.SelectionStart);
            DialogResult = DialogResult.OK;
            Close();
            e.Handled = true;
        }
    }
}
