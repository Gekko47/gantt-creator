using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// WinForms modal dialog for R5.6 theme settings (Colours / Heights and sizes / Grid and bands tabs).
/// Runs on the Excel main STA thread via ShowDialog; fully offline.
/// </summary>
/// <summary>
/// Snapshot of one custom style's full property state: the colours, the fill
/// pattern, the sizes, and the label vocabulary.
/// </summary>
/// <remarks>
/// Every member mirrors one column of <c>tblGanttStyles</c>, so the dialog
/// edits an object rather than a bag of loosely related options. The sizes
/// and the label vocabulary are carried here because they belong to the
/// style, not to the chart: a custom style used by a Custom Activity row and
/// the same style used by a Custom Milestone row are one object with one set
/// of properties.
/// </remarks>
internal sealed class CustomStyleState(
    string fill,
    string stroke,
    string text,
    GanttHatchPattern hatchPattern = GanttHatchPattern.None,
    double standardOutlinePt = 0,
    double activityHeightPt = 0,
    double milestoneSizePt = 0,
    GanttLabelPosition defaultLabelPosition = GanttLabelPosition.Auto,
    IReadOnlySet<GanttLabelPosition>? allowedLabelPositions = null,
    GanttStrokeDash strokeDash = GanttStrokeDash.Solid)
{
    public string Fill { get; } = fill;
    public string Stroke { get; } = stroke;
    public string Text { get; } = text;
    public GanttHatchPattern HatchPattern { get; } = hatchPattern;
    public double StandardOutlinePt { get; } = standardOutlinePt;
    public double ActivityHeightPt { get; } = activityHeightPt;
    public double MilestoneSizePt { get; } = milestoneSizePt;
    public GanttLabelPosition DefaultLabelPosition { get; } = defaultLabelPosition;
    public IReadOnlySet<GanttLabelPosition> AllowedLabelPositions { get; } =
        allowedLabelPositions ?? new HashSet<GanttLabelPosition> { GanttLabelPosition.Auto };
    public GanttStrokeDash StrokeDash { get; } = strokeDash;
}

/// <summary>
/// One built-in style's pending colour values, held while the master–detail
/// view shows a different style.
/// </summary>
/// <param name="Fill">The fill colour, or empty when the style has no fill.</param>
/// <param name="Stroke">The outline colour, or empty when the style has no outline.</param>
/// <param name="Text">The label text colour, or empty.</param>
internal sealed record EntityStyleColours(string Fill, string Stroke, string Text);

internal sealed partial class ThemeSettingsForm : Form
{
    private ThemeEditRequest? _validatedRequest;
    private readonly IReadOnlyList<CustomStyleDef> _originalCustomDefs = [];

    private readonly TabControl _tabControl = new()
    {
        Dock = DockStyle.Fill,
        Padding = new Point(6, 3),
    };

    private readonly Button _okButton = new()
    {
        Text = "OK",
        DialogResult = DialogResult.OK,
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
    };

    private readonly Button _cancelButton = new()
    {
        Text = "Cancel",
        DialogResult = DialogResult.Cancel,
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
    };

    private readonly Button _resetButton = new()
    {
        Text = "Reset to default",
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
    };

    private readonly Label _warningLabel = new()
    {
        Dock = DockStyle.Top,
        Height = 24,
        ForeColor = Color.FromArgb(192, 0, 0),
        TextAlign = ContentAlignment.MiddleLeft,
        Visible = false,
    };

    private readonly Dictionary<string, TextBox> _colorTextBoxes = [];

    /// <summary>
    /// The colour swatch beside each editable colour, keyed exactly like
    /// <see cref="_colorTextBoxes"/>. The swatch used to be found by walking
    /// the form's direct children, which never matched: each swatch lives
    /// inside a group box, so the swatch kept showing the code-owned default
    /// while the text box beside it showed the workbook's colour.
    /// </summary>
    private readonly Dictionary<string, Panel> _colorSwatches = [];

    /// <summary>
    /// The chart-chrome colour swatches, keyed by the setting key the colour
    /// is stored under. They have their own dictionary for the same reason as
    /// <see cref="_colorSwatches"/>.
    /// </summary>
    private readonly Dictionary<string, Panel> _chromeColourSwatches = [];

    /// <summary>
    /// The chart-chrome colour text boxes, keyed by the setting key the colour
    /// is stored under. They live on the Chart tab, because the colours belong
    /// to the chart as a whole rather than to any entity's style.
    /// </summary>
    private readonly Dictionary<string, TextBox> _chromeColourTextBoxes = [];
    private readonly Dictionary<string, Button> _colorPickerButtons = [];
    private readonly Dictionary<string, NumericUpDown> _metricNumericUpDowns = [];
    private readonly Dictionary<string, Label> _metricLabels = [];

    // Custom styles tab controls
    // Entity styles tab controls (master-detail: a list of styles, and the
    // selected style's properties beside it).
    private readonly ListView _entityStyleList = new();
    private readonly Panel _entityStyleDetailPanel = new();
    private readonly Label _entityStyleTitleLabel = new();
    private readonly TextBox _entityFillText = new();
    private readonly TextBox _entityStrokeText = new();
    private readonly TextBox _entityTextText = new();
    private readonly Button _entityFillPicker = new();
    private readonly Button _entityStrokePicker = new();
    private readonly Button _entityTextPicker = new();
    private readonly Panel _entityFillSwatch = new();
    private readonly Panel _entityStrokeSwatch = new();
    private readonly Panel _entityTextSwatch = new();
    private readonly Label _entityPatternLabel = new();
    private readonly Label _entityDashLabel = new();
    private readonly ComboBox _entityDashCombo = new();
    private readonly Label _entityOutlineWidthLabel = new();
    private readonly Label _entitySizeLabel = new();
    private readonly Label _entityLabelPositionLabel = new();
    private readonly NumericUpDown _entityLineWidthNumeric = new();

    /// <summary>
    /// The pending colour values for each built-in style the user has edited
    /// or visited, keyed by the style key. The master-detail view shows one
    /// style at a time, so the values for the styles that are not on screen
    /// live here rather than in a text box.
    /// </summary>
    private readonly Dictionary<string, EntityStyleColours> _entityStyleColours = [];

    /// <summary>
    /// The style key whose properties the detail panel is currently showing.
    /// </summary>
    private string? _selectedEntityStyleKey;

    private const string _customStyleKeyLabel = "Style key:";
    private const string _customStyleEmptyHint = "No custom styles defined yet. Enter a key and click New custom style.";

    private readonly ListBox _customStyleList = new()
    {
        Size = new Size(200, 180),
        IntegralHeight = false,
    };
    private readonly Label _customStyleListLabel = new()
    {
        Text = "Custom styles:",
        Size = new Size(100, 20),
    };
    private readonly Label _customStyleListHint = new()
    {
        Text = _customStyleEmptyHint,
        Size = new Size(190, 20),
        TextAlign = ContentAlignment.MiddleLeft,
        ForeColor = SystemColors.GrayText,
        Visible = false,
    };
    private readonly TextBox _customStyleKeyText = new()
    {
        Size = new Size(140, 22),
        CharacterCasing = CharacterCasing.Upper,
        MaxLength = 32,
    };
    private readonly TextBox _customStyleFillText = new()
    {
        Size = new Size(100, 22),
        CharacterCasing = CharacterCasing.Upper,
        MaxLength = 9,
    };
    private readonly TextBox _customStyleStrokeText = new()
    {
        Size = new Size(100, 22),
        CharacterCasing = CharacterCasing.Upper,
        MaxLength = 9,
    };
    private readonly TextBox _customStyleTextText = new()
    {
        Size = new Size(100, 22),
        CharacterCasing = CharacterCasing.Upper,
        MaxLength = 9,
    };
    private readonly Button _customFillColorButton = new()
    {
        Text = "\u2026",
        Size = new Size(26, 24),
    };
    private readonly Button _customStrokeColorButton = new()
    {
        Text = "\u2026",
        Size = new Size(26, 24),
    };
    private readonly Button _customTextColorButton = new()
    {
        Text = "\u2026",
        Size = new Size(26, 24),
    };
    private readonly Panel _customFillSwatch = new()
    {
        Size = new Size(24, 22),
        BorderStyle = BorderStyle.FixedSingle,
        BackColor = SystemColors.Control,
    };
    private readonly Panel _customStrokeSwatch = new()
    {
        Size = new Size(24, 22),
        BorderStyle = BorderStyle.FixedSingle,
        BackColor = SystemColors.Control,
    };
    private readonly Panel _customTextSwatch = new()
    {
        Size = new Size(24, 22),
        BorderStyle = BorderStyle.FixedSingle,
        BackColor = SystemColors.Control,
    };
    private readonly Button _customNewButton = new()
    {
        Text = "New custom style",
        AutoSize = true,
    };
    private readonly TextBox _customNewKeyText = new()
    {
        Size = new Size(120, 22),
        CharacterCasing = CharacterCasing.Upper,
        MaxLength = 32,
    };
    private readonly Button _customDeleteButton = new()
    {
        Text = "Delete",
        AutoSize = true,
        Enabled = false,
    };

    // The fill pattern, the outline width, the two sizes, and the label
    // vocabulary are properties of the STYLE, so they sit in the style editor
    // beside the colours rather than in the chart-wide metric tabs.
    private readonly ComboBox _customHatchPatternCombo = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Size = new Size(140, 22),
    };
    private readonly NumericUpDown _customOutlineWidthNumeric = new()
    {
        Size = new Size(80, 22),
        DecimalPlaces = 2,
        Increment = 0.25m,
        Minimum = 0m,
    };
    private readonly NumericUpDown _customActivityHeightNumeric = new()
    {
        Size = new Size(80, 22),
        DecimalPlaces = 2,
        Increment = 0.25m,
        Minimum = 0m,
    };
    private readonly NumericUpDown _customMilestoneSizeNumeric = new()
    {
        Size = new Size(80, 22),
        DecimalPlaces = 2,
        Increment = 0.25m,
        Minimum = 0m,
    };
    private readonly ComboBox _customDefaultLabelCombo = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Size = new Size(140, 22),
    };
    private readonly ComboBox _customStrokeDashCombo = new()
    {
        DropDownStyle = ComboBoxStyle.DropDownList,
        Size = new Size(140, 22),
    };
    private readonly CheckedListBox _customAllowedLabelList = new()
    {
        CheckOnClick = true,
        IntegralHeight = false,
        Size = new Size(140, 74),
    };
    private readonly GroupBox _customStyleListGroup = new()
    {
        Text = "Custom styles",
        Size = new Size(220, 260),
    };
    private readonly GroupBox _customStyleEditGroup = new()
    {
        Text = "Edit selected style",
        Size = new Size(320, 260),
        Enabled = false,
    };

    private string? _selectedCustomStyleKey;
    private readonly Dictionary<string, CustomStyleState> _customStyleEdits = [];

    /// <summary>
    /// Suppresses the editor's change handlers while the editor is being
    /// populated from the stored state. Without it, loading a style would fire
    /// the change handlers and mark the style as edited — the dialog would
    /// then emit an update for a style the user never touched.
    /// </summary>
    private bool _suppressCustomStyleSync;

    private readonly ColorDialog _colorDialog = new()
    {
        AllowFullOpen = true,
        SolidColorOnly = true,
        AnyColor = true,
    };

    /// <summary>
    /// Initializes the dialog with the current theme values.
    /// </summary>
    /// <param name="request">The current theme edit request (pre-populated with workbook values).</param>
    internal ThemeSettingsForm(ThemeEditRequest request)
    {
        _originalCustomDefs = request.CustomStyleDefs ?? [];
        InitializeComponent();
        PopulateFromRequest(request);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _colorDialog.Dispose();
            _warningLabel.Dispose();
            _cancelButton.Dispose();
            _resetButton.Dispose();
            _okButton.Dispose();
            _tabControl.Dispose();
            _customStyleList.Dispose();
            _customStyleListLabel.Dispose();
            _customStyleListHint.Dispose();
            _customStyleKeyText.Dispose();
            _customStyleFillText.Dispose();
            _customStyleStrokeText.Dispose();
            _customStyleTextText.Dispose();
            _customFillColorButton.Dispose();
            _customStrokeColorButton.Dispose();
            _customTextColorButton.Dispose();
            _customFillSwatch.Dispose();
            _customStrokeSwatch.Dispose();
            _customTextSwatch.Dispose();
            _customNewButton.Dispose();
            _customNewKeyText.Dispose();
            _customDeleteButton.Dispose();
            _customStyleListGroup.Dispose();
            _customStyleEditGroup.Dispose();
            _customHatchPatternCombo.Dispose();
            _customOutlineWidthNumeric.Dispose();
            _customActivityHeightNumeric.Dispose();
            _customMilestoneSizeNumeric.Dispose();
            _customDefaultLabelCombo.Dispose();
            _customStrokeDashCombo.Dispose();
            _customAllowedLabelList.Dispose();
            _entityStyleList.Dispose();
            _entityStyleDetailPanel.Dispose();
            _entityStyleTitleLabel.Dispose();
            _entityFillText.Dispose();
            _entityStrokeText.Dispose();
            _entityTextText.Dispose();
            _entityFillPicker.Dispose();
            _entityStrokePicker.Dispose();
            _entityTextPicker.Dispose();
            _entityFillSwatch.Dispose();
            _entityStrokeSwatch.Dispose();
            _entityTextSwatch.Dispose();
            _entityPatternLabel.Dispose();
            _entityDashLabel.Dispose();
            _entityDashCombo.Dispose();
            _entityOutlineWidthLabel.Dispose();
            _entitySizeLabel.Dispose();
            _entityLabelPositionLabel.Dispose();
            _entityLineWidthNumeric.Dispose();
        }
        base.Dispose(disposing);
    }

    /// <summary>
    /// Gets the validated request if the user clicked OK, or null if cancelled/invalid.
    /// </summary>
    internal ThemeEditRequest? GetValidatedRequest() => _validatedRequest;

    /// <summary>
    /// Test hook: sets the validated request without running the dialog's OK
    /// handler. Used by RibbonStateServiceTests so ShowDialog() can be
    /// substituted with a non-blocking result.
    /// </summary>
    internal void SetValidatedRequestForTesting(ThemeEditRequest? request) => _validatedRequest = request;

    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void InitializeComponent()
    {
        Text = "Theme settings";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(720, 560);
        MinimumSize = new Size(720, 560);
        AutoScaleMode = AutoScaleMode.Dpi;

        Controls.Add(_tabControl);

        CreateColoursTab();
        CreateChartTab();
        CreateCustomStylesTab();

        // Bottom buttons + warning label
        Panel bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            Padding = new Padding(8),
        };
        Controls.Add(bottomPanel);

        _warningLabel.Font = new Font(Font.FontFamily, 8.25f, FontStyle.Regular);
        bottomPanel.Controls.Add(_warningLabel);

        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Padding = new Padding(0, 4, 0, 0),
        };
        bottomPanel.Controls.Add(buttonPanel);

        buttonPanel.Controls.Add(_cancelButton);
        _resetButton.Click += ResetButton_Click;
        buttonPanel.Controls.Add(_resetButton);
        _okButton.Click += OkButton_Click;
        buttonPanel.Controls.Add(_okButton);

        AcceptButton = _okButton;
        CancelButton = _cancelButton;
    }

    /// <summary>
    /// Builds the detail panel: the selected style's properties, grouped by
    /// concern rather than by table column.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void BuildEntityStyleDetailPanel()
    {
        _entityStyleTitleLabel.Dock = DockStyle.Top;
        _entityStyleTitleLabel.Height = 26;
        _entityStyleTitleLabel.Font = new Font(Font, FontStyle.Bold);
        _entityStyleTitleLabel.TextAlign = ContentAlignment.MiddleLeft;
        _entityStyleDetailPanel.Controls.Add(_entityStyleTitleLabel);

        FlowLayoutPanel body = new()
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            AutoScroll = true,
            WrapContents = false,
            Padding = new Padding(4, 4, 4, 4),
        };
        _entityStyleDetailPanel.Controls.Add(body);

        body.Controls.Add(BuildEntityFillGroup());
        body.Controls.Add(BuildEntityOutlineGroup());
        body.Controls.Add(BuildEntityLabelGroup());
        body.Controls.Add(BuildEntitySizeGroup());
    }

    /// <summary>
    /// The Fill group: the fill colour and the fill pattern.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntityFillGroup()
    {
        GroupBox group = new() { Text = "Fill", Width = 380, AutoSize = true };
        TableLayoutPanel table = NewEntityPropertyTable();
        group.Controls.Add(table);

        _entityFillText.Name = "entityFill";
        _entityFillText.MaxLength = 9;
        _entityFillText.CharacterCasing = CharacterCasing.Upper;
        _entityFillText.Width = 110;
        _entityFillText.Leave += (_, _) => ValidateColourTextBox(_entityFillText);
        _entityFillText.TextChanged += (_, _) => CaptureSelectedEntityColours();

        _entityFillPicker.Text = "\u2026";
        _entityFillPicker.Width = 26;
        _entityFillPicker.Tag = "entityFill";
        _entityFillPicker.Click += ColorPickerButton_Click;

        _entityFillSwatch.Size = new Size(24, 22);
        _entityFillSwatch.BorderStyle = BorderStyle.FixedSingle;
        _entityFillSwatch.Tag = "entityFill";

        table.Controls.Add(new Label { Text = "Colour", AutoSize = true }, 0, 0);
        table.Controls.Add(_entityFillText, 1, 0);
        table.Controls.Add(_entityFillPicker, 2, 0);
        table.Controls.Add(_entityFillSwatch, 3, 0);

        _entityPatternLabel.AutoSize = true;
        _entityPatternLabel.ForeColor = SystemColors.GrayText;
        table.Controls.Add(new Label { Text = "Pattern", AutoSize = true }, 0, 1);
        table.Controls.Add(_entityPatternLabel, 1, 1);
        table.SetColumnSpan(_entityPatternLabel, 3);

        _colorTextBoxes["entityFill"] = _entityFillText;
        _colorPickerButtons["entityFill"] = _entityFillPicker;
        _colorSwatches["entityFill"] = _entityFillSwatch;

        return group;
    }

    /// <summary>
    /// Creates the four-column table every entity property group uses:
    /// label, value, picker button, swatch.
    /// </summary>
    private static TableLayoutPanel NewEntityPropertyTable()
    {
        TableLayoutPanel table = new()
        {
            ColumnCount = 4,
            AutoSize = true,
            Dock = DockStyle.Top,
        };
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 30));
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        return table;
    }
    /// <summary>
    /// The Outline group: the outline colour, its width, and its line type.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntityOutlineGroup()
    {
        GroupBox group = new() { Text = "Outline", Width = 380, AutoSize = true };
        TableLayoutPanel table = NewEntityPropertyTable();
        group.Controls.Add(table);

        _entityStrokeText.Name = "entityStroke";
        _entityStrokeText.MaxLength = 9;
        _entityStrokeText.CharacterCasing = CharacterCasing.Upper;
        _entityStrokeText.Width = 110;
        _entityStrokeText.Leave += (_, _) => ValidateColourTextBox(_entityStrokeText);
        _entityStrokeText.TextChanged += (_, _) => CaptureSelectedEntityColours();

        _entityStrokePicker.Text = "\u2026";
        _entityStrokePicker.Width = 26;
        _entityStrokePicker.Tag = "entityStroke";
        _entityStrokePicker.Click += ColorPickerButton_Click;

        _entityStrokeSwatch.Size = new Size(24, 22);
        _entityStrokeSwatch.BorderStyle = BorderStyle.FixedSingle;
        _entityStrokeSwatch.Tag = "entityStroke";

        table.Controls.Add(new Label { Text = "Colour", AutoSize = true }, 0, 0);
        table.Controls.Add(_entityStrokeText, 1, 0);
        table.Controls.Add(_entityStrokePicker, 2, 0);
        table.Controls.Add(_entityStrokeSwatch, 3, 0);

        _entityOutlineWidthLabel.AutoSize = true;
        _entityOutlineWidthLabel.ForeColor = SystemColors.GrayText;
        table.Controls.Add(new Label { Text = "Width", AutoSize = true }, 0, 1);
        table.Controls.Add(_entityOutlineWidthLabel, 1, 1);
        table.SetColumnSpan(_entityOutlineWidthLabel, 3);

        _entityDashLabel.AutoSize = true;
        _entityDashLabel.ForeColor = SystemColors.GrayText;
        _entityDashCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        _entityDashCombo.Width = 140;

        table.Controls.Add(new Label { Text = "Line type", AutoSize = true }, 0, 2);
        table.Controls.Add(_entityDashCombo, 1, 2);
        table.Controls.Add(_entityDashLabel, 2, 2);
        table.SetColumnSpan(_entityDashLabel, 2);

        _colorTextBoxes["entityStroke"] = _entityStrokeText;
        _colorPickerButtons["entityStroke"] = _entityStrokePicker;
        _colorSwatches["entityStroke"] = _entityStrokeSwatch;

        return group;
    }

    /// <summary>
    /// The Label group: the label colour and the label vocabulary.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntityLabelGroup()
    {
        GroupBox group = new() { Text = "Label", Width = 380, AutoSize = true };
        TableLayoutPanel table = NewEntityPropertyTable();
        group.Controls.Add(table);

        _entityTextText.Name = "entityText";
        _entityTextText.MaxLength = 9;
        _entityTextText.CharacterCasing = CharacterCasing.Upper;
        _entityTextText.Width = 110;
        _entityTextText.Leave += (_, _) => ValidateColourTextBox(_entityTextText);
        _entityTextText.TextChanged += (_, _) => CaptureSelectedEntityColours();

        _entityTextPicker.Text = "\u2026";
        _entityTextPicker.Width = 26;
        _entityTextPicker.Tag = "entityText";
        _entityTextPicker.Click += ColorPickerButton_Click;

        _entityTextSwatch.Size = new Size(24, 22);
        _entityTextSwatch.BorderStyle = BorderStyle.FixedSingle;
        _entityTextSwatch.Tag = "entityText";

        table.Controls.Add(new Label { Text = "Colour", AutoSize = true }, 0, 0);
        table.Controls.Add(_entityTextText, 1, 0);
        table.Controls.Add(_entityTextPicker, 2, 0);
        table.Controls.Add(_entityTextSwatch, 3, 0);

        _entityLabelPositionLabel.AutoSize = true;
        _entityLabelPositionLabel.ForeColor = SystemColors.GrayText;
        table.Controls.Add(new Label { Text = "Position", AutoSize = true }, 0, 1);
        table.Controls.Add(_entityLabelPositionLabel, 1, 1);
        table.SetColumnSpan(_entityLabelPositionLabel, 3);

        _colorTextBoxes["entityText"] = _entityTextText;
        _colorPickerButtons["entityText"] = _entityTextPicker;
        _colorSwatches["entityText"] = _entityTextSwatch;

        return group;
    }

    /// <summary>
    /// The Size group: the bar height or milestone size, and the delineator's
    /// own line width.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntitySizeGroup()
    {
        GroupBox group = new() { Text = "Size", Width = 380, AutoSize = true };
        TableLayoutPanel table = new()
        {
            ColumnCount = 2,
            AutoSize = true,
            Dock = DockStyle.Top,
        };
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        group.Controls.Add(table);

        _entitySizeLabel.AutoSize = true;
        _entitySizeLabel.ForeColor = SystemColors.GrayText;
        table.Controls.Add(new Label { Text = "Entity", AutoSize = true }, 0, 0);
        table.Controls.Add(_entitySizeLabel, 1, 0);

        // The delineator's line width is the one built-in size the user may
        // change: it is the chart-wide DelineatorLinePt metric, which no
        // other entity reads.
        _entityLineWidthNumeric.Minimum = 0;
        _entityLineWidthNumeric.Maximum = 20;
        _entityLineWidthNumeric.DecimalPlaces = 2;
        _entityLineWidthNumeric.Increment = 0.25m;
        _entityLineWidthNumeric.Width = 80;

        table.Controls.Add(new Label { Text = "Line width", AutoSize = true }, 0, 1);
        table.Controls.Add(_entityLineWidthNumeric, 1, 1);

        return group;
    }

    /// <summary>
    /// Shows the selected style's properties in the detail panel.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void EntityStyleList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        CaptureSelectedEntityColours();

        if (_entityStyleList.SelectedItems.Count == 0
            || _entityStyleList.SelectedItems[0].Tag is not string styleKey)
        {
            _selectedEntityStyleKey = null;
            _entityStyleTitleLabel.Text = "Select a style to see its properties.";
            return;
        }

        ShowEntityStyle(styleKey);
    }

    /// <summary>
    /// Test hook: shows a style's properties without the list view's native
    /// selection state, which a unit test cannot create.
    /// </summary>
    /// <param name="styleKey">The built-in style key.</param>
    internal void ShowEntityStyleForTesting(string styleKey) => ShowEntityStyle(styleKey);

    /// <summary>
    /// Shows one style's properties in the detail panel.
    /// </summary>
    /// <param name="styleKey">The built-in style key.</param>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void ShowEntityStyle(string styleKey)
    {
        _selectedEntityStyleKey = styleKey;
        GanttStylePreset preset = GanttCatalogues.GetPreset(styleKey);
        _entityStyleTitleLabel.Text = preset.DisplayName;
        EntityStyleColours colours = _entityStyleColours.TryGetValue(styleKey, out EntityStyleColours? known)
            ? known
            : new EntityStyleColours(preset.FillColour, preset.StrokeColour, preset.TextColour);

        _entityFillText.Text = colours.Fill;
        _entityStrokeText.Text = colours.Stroke;
        _entityTextText.Text = colours.Text;
        UpdateSwatch("entityFill", ParseColourOrDefault(colours.Fill));
        UpdateSwatch("entityStroke", ParseColourOrDefault(colours.Stroke));
        UpdateSwatch("entityText", ParseColourOrDefault(colours.Text));

        _entityPatternLabel.Text = DescribeHatchPattern(preset.HatchPattern);

        // The general dash, the outline width, and the bar height are
        // code-owned for a built-in style: the reader exact-matches those
        // columns, so an editable control would offer a change the workbook
        // could not keep. The delineator's line type is the exception.
        bool isDelineator = string.Equals(styleKey, "DefaultDelineator", StringComparison.Ordinal);
        _entityDashCombo.Visible = isDelineator;
        _entityDashLabel.Visible = !isDelineator;
        _entityDashLabel.Text = DescribeStrokeDash(preset.StrokeDash);
        if (isDelineator)
        {
            PopulateDashVocabulary(_entityDashCombo);
            SelectDash(_entityDashCombo, preset.StrokeDash);
        }

        _entityOutlineWidthLabel.Text = preset.StandardOutlinePt > 0
            ? $"{preset.StandardOutlinePt:0.##} pt (fixed by the built-in style)"
            : "Not applicable";

        _entitySizeLabel.Text = DescribeEntitySize(preset);

        // The delineator's line width is the only size the user may change.
        _entityLineWidthNumeric.Enabled = isDelineator;
        _entityLineWidthNumeric.Visible = isDelineator;
    }

    /// <summary>
    /// Describes an entity's size, which is a bar height for a span and a
    /// diamond size for a milestone.
    /// </summary>
    private static string DescribeEntitySize(GanttStylePreset preset) =>
        preset.ActivityHeightPt > 0
            ? $"Bar height {preset.ActivityHeightPt:0.##} pt (fixed by the built-in style)"
            : preset.MilestoneSizePt > 0
                ? $"Milestone size {preset.MilestoneSizePt:0.##} pt (fixed by the built-in style)"
                : "Not applicable";

    /// <summary>
    /// Captures the detail panel's colour values for the selected style.
    /// </summary>
    private void CaptureSelectedEntityColours()
    {
        if (_selectedEntityStyleKey is not { } styleKey)
        {
            return;
        }

        _entityStyleColours[styleKey] = new EntityStyleColours(
            _entityFillText.Text,
            _entityStrokeText.Text,
            _entityTextText.Text);
    }

    /// <summary>
    /// Returns the dash the delineator's line-type combo currently shows.
    /// </summary>
    private GanttStrokeDash? SelectedEntityDash() =>
        _entityDashCombo.SelectedItem is StrokeDashItem item ? item.Dash : null;

    /// <summary>
    /// Fills a dash combo with the closed dash vocabulary.
    /// </summary>
    private static void PopulateDashVocabulary(ComboBox combo)
    {
        combo.Items.Clear();
        foreach (GanttStrokeDash dash in Enum.GetValues<GanttStrokeDash>())
        {
            _ = combo.Items.Add(new StrokeDashItem(dash, DescribeStrokeDash(dash)));
        }
    }

    /// <summary>
    /// Selects the dash a combo shows for a given pattern.
    /// </summary>
    private static void SelectDash(ComboBox combo, GanttStrokeDash dash)
    {
        for (int index = 0; index < combo.Items.Count; index++)
        {
            if (combo.Items[index] is StrokeDashItem item && item.Dash == dash)
            {
                combo.SelectedIndex = index;
                return;
            }
        }
    }

    /// <summary>
    /// The entity styles tab: a master–detail view. The left list holds the
    /// built-in styles grouped by their family, and the right panel holds
    /// every property of the selected style.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void CreateColoursTab()
    {
        TabPage tab = new TabPage("Entity styles")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        SplitContainer split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            SplitterDistance = 250,
            FixedPanel = FixedPanel.Panel1,
        };
        tab.Controls.Add(split);

        // The master: one list, grouped by family. The group heading comes
        // from the code-owned family authority, so a style can never appear
        // under a second heading and no heading is emitted twice.
        _entityStyleList.View = View.Details;
        _entityStyleList.FullRowSelect = true;
        _entityStyleList.HideSelection = false;
        _entityStyleList.HeaderStyle = ColumnHeaderStyle.None;
        _entityStyleList.Dock = DockStyle.Fill;
        _entityStyleList.SelectedIndexChanged += EntityStyleList_SelectedIndexChanged;
        split.Panel1.Controls.Add(_entityStyleList);

        // The detail: the selected style's properties.
        _entityStyleDetailPanel.Dock = DockStyle.Fill;
        _entityStyleDetailPanel.AutoScroll = true;
        _entityStyleDetailPanel.Padding = new Padding(8);
        split.Panel2.Controls.Add(_entityStyleDetailPanel);

        BuildEntityStyleDetailPanel();

        Label title = new()
        {
            Text = "Select a style to see its properties.",
            Dock = DockStyle.Top,
            Height = 24,
            ForeColor = SystemColors.GrayText,
            TextAlign = ContentAlignment.MiddleLeft,
        };
        _entityStyleDetailPanel.Controls.Add(title);

        foreach (GanttEntityFamily family in GanttStyleFamilies.OrderedFamilies)
        {
            GanttStylePreset[] members =
            [
                .. GanttCatalogues.StylePresets
                    .Where(preset => GanttStyleFamilies.For(preset.StyleKey) == family)
                    .Where(preset =>
                        preset.ColourCapability.HasFlag(EntityColourCapability.Fill)
                        || preset.ColourCapability.HasFlag(EntityColourCapability.Stroke)
                        || preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined)),
            ];


            if (members.Length == 0)
            {
                continue;
            }

            ListViewGroup group = new(GanttStyleFamilies.DisplayNameFor(family));
            _ = _entityStyleList.Groups.Add(group);

            foreach (GanttStylePreset preset in members)
            {
                ListViewItem item = new(preset.DisplayName) { Group = group, Tag = preset.StyleKey };
                _ = _entityStyleList.Items.Add(item);
            }
        }
    }

    private void ColorPickerButton_Click(object? sender, EventArgs e)
    {
        if (sender is Button btn && btn.Tag is string key &&
            (_colorTextBoxes.TryGetValue(key, out TextBox? txt)
                || _chromeColourTextBoxes.TryGetValue(key, out txt)))
        {
            if (ColourHex.TryParse(txt.Text, out ColourHex? initial))
            {
                _colorDialog.Color = Color.FromArgb((int)initial!.ARGB);
            }

            if (_colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                Color color = _colorDialog.Color;
                string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
                txt.Text = hex;
                UpdateSwatch(key, color);
            }
        }
    }

    private void ValidateColourTextBox(TextBox txt)
    {
        string text = txt.Text.Trim();
        if (string.IsNullOrEmpty(text))
        {
            txt.BackColor = SystemColors.Window;
            return;
        }

        if (ColourHex.TryParse(text, out _))
        {
            txt.BackColor = SystemColors.Window;
            UpdateSwatch(txt.Name, ColorTranslator.FromHtml(text));
        }
        else
        {
            txt.BackColor = Color.FromArgb(255, 230, 230);
        }
    }

    private void UpdateSwatch(string key, Color color)
    {
        // The swatch is found through the dictionary, not by walking the
        // control tree: the walk only ever saw the form's DIRECT children, so
        // a swatch inside a group box was never found and the swatch kept
        // showing the code-owned default while the text box beside it showed
        // the workbook's colour.
        if (_colorSwatches.TryGetValue(key, out Panel? swatch))
        {
            swatch.BackColor = color;
            return;
        }

        if (_chromeColourSwatches.TryGetValue(key, out Panel? chromeSwatch))
        {
            chromeSwatch.BackColor = color;
        }
    }

    private static Color ParseColourOrDefault(string value)
    {
        return ColourHex.TryParse(value, out ColourHex? c) && c is not null
            ? Color.FromArgb((int)c.ARGB)
            : SystemColors.Control;
    }

    /// <summary>
    /// Creates the chart-wide tab. The options are grouped by the chart object
    /// they belong to, not by the metric token's name, so the user can see
    /// which knob changes which part of the chart.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void CreateChartTab()
    {
        TabPage tab = new TabPage("Chart")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        tab.Controls.Add(panel);

        // The chart-wide knobs only. Two earlier groups are gone: "Entity
        // defaults" carried the bar height and the outline width, which the
        // scene resolves from the STYLE row rather than from these metrics, so
        // the controls promised a change the chart never rendered; and "Hatch
        // spacing" carried the pitch and the line width, which the host's
        // MsoPatternType cannot express at all.
        int y = 8;
        y = AddMetricGroup(panel, "Rows and lanes", y, "GanttRowHeightPt", "SplitterHeightPt", "SpacerHeightPt");
        y = AddMetricGroup(panel, "Grid", y, "GridLinePt", "MajorBoundaryPt");
        y = AddMetricGroup(panel, "Milestones", y, "MilestoneSizePt");
        y = AddMetricGroup(panel, "Delineator", y, "DelineatorLinePt");
        _ = AddChromeColourGroup(panel, y);
    }

    /// <summary>
    /// Adds a titled group of metric controls to the chart tab.
    /// </summary>
    /// <param name="panel">The tab's panel.</param>
    /// <param name="title">The group title.</param>
    /// <param name="top">The group's top offset.</param>
    /// <param name="metricKeys">The metric tokens in the group, in display order.</param>
    /// <returns>The next free top offset.</returns>
    private int AddMetricGroup(Panel panel, string title, int top, params string[] metricKeys)
    {
        List<(string Key, GanttMetricToken Token)> tokens = [];
        foreach (string key in metricKeys)
        {
            GanttMetricToken? token = GanttCatalogues.Metrics.FirstOrDefault(candidate =>
                string.Equals(candidate.Name, key, StringComparison.Ordinal));
            if (token is not null)
            {
                tokens.Add((key, token));
            }
        }

        if (tokens.Count == 0)
        {
            return top;
        }

        GroupBox group = new GroupBox
        {
            Text = title,
            Location = new Point(8, top),
            Size = new Size(660, (tokens.Count * 26) + 20),
        };
        panel.Controls.Add(group);

        int innerY = 24;
        foreach ((string key, GanttMetricToken token) in tokens)
        {
            Label lbl = new()
            {
                Text = $"{DescribeMetric(key)} ({token.Minimum} to {token.Maximum} pt)",
                Location = new Point(12, innerY),
                Size = new Size(300, 20),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            group.Controls.Add(lbl);

            NumericUpDown nud = new()
            {
                Name = key,
                Location = new Point(460, innerY - 2),
                Size = new Size(90, 22),
                Minimum = (decimal)token.Minimum,
                Maximum = (decimal)token.Maximum,
                DecimalPlaces = 2,
                Increment = 0.25m,
            };
            group.Controls.Add(nud);
            _metricNumericUpDowns[key] = nud;
            _metricLabels[key] = lbl;

            innerY += 26;
        }

        return top + group.Height + 10;
    }

    /// <summary>
    /// Adds the chart-chrome colour group: the colours of the chart frame the
    /// entities sit on.
    /// </summary>
    /// <param name="panel">The tab's panel.</param>
    /// <param name="top">The group's top offset.</param>
    /// <returns>The next free top offset.</returns>
    /// <remarks>
    /// The group is object-first like the rest of the dialog: each row is one
    /// chrome role with its own colour, rather than a column of colours with
    /// no object named. The roles are the entity guide's own chart-chrome
    /// tokens, so the label the user reads is the token the contract names.
    /// </remarks>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private int AddChromeColourGroup(Panel panel, int top)
    {
        const int RowHeight = 28;
        GroupBox group = new()
        {
            Text = "Chart appearance",
            Location = new Point(8, top),
            Size = new Size(660, (GanttChartChrome.Roles.Count * RowHeight) + 20),
        };
        panel.Controls.Add(group);

        int innerY = 24;
        foreach (GanttChartChromeRole role in GanttChartChrome.Roles)
        {
            string key = GanttChartChrome.SettingKeyFor(role);
            string label = GanttChartChrome.DisplayLabelFor(role);

            Label lbl = new()
            {
                Text = label,
                Location = new Point(12, innerY + 3),
                Size = new Size(120, 20),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            group.Controls.Add(lbl);

            TextBox txt = new()
            {
                Name = key,
                Location = new Point(140, innerY),
                Size = new Size(100, 22),
                Text = GanttChartChrome.DefaultHexFor(role),
                CharacterCasing = CharacterCasing.Upper,
                MaxLength = 9,
            };
            txt.Leave += (_, _) => ValidateColourTextBox(txt);
            group.Controls.Add(txt);
            _chromeColourTextBoxes[key] = txt;

            Button picker = new()
            {
                Text = "\u2026",
                Location = new Point(246, innerY - 1),
                Size = new Size(26, 24),
                Tag = key,
            };
            picker.Click += ColorPickerButton_Click;
            group.Controls.Add(picker);

            Panel swatch = new()
            {
                Location = new Point(280, innerY),
                Size = new Size(24, 22),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = ParseColourOrDefault(txt.Text),
                Tag = key,
            };
            group.Controls.Add(swatch);
            _chromeColourSwatches[key] = swatch;

            innerY += RowHeight;
        }

        return top + group.Height + 10;
    }

    /// <summary>
    /// Finds the chrome role a setting key names.
    /// </summary>
    /// <param name="settingKey">The setting key.</param>
    /// <returns>The chrome role.</returns>
    /// <exception cref="ArgumentOutOfRangeException">Thrown when the key names no role.</exception>
    private static GanttChartChromeRole ChromeRoleFor(string settingKey)
    {
        foreach (GanttChartChromeRole role in GanttChartChrome.Roles)
        {
            if (string.Equals(GanttChartChrome.SettingKeyFor(role), settingKey, StringComparison.Ordinal))
            {
                return role;
            }
        }

        throw new ArgumentOutOfRangeException(nameof(settingKey), settingKey, "The key names no chart chrome role.");
    }

    /// <summary>
    /// Returns the user-facing name of a metric token.
    /// </summary>
    /// <param name="key">The metric token name.</param>
    /// <returns>The display name.</returns>
    private static string DescribeMetric(string key) => key switch
    {
        "GanttRowHeightPt" => "Row height",
        "SplitterHeightPt" => "Splitter height",
        "SpacerHeightPt" => "Spacer height",
        "ActivityHeightPt" => "Bar height",
        "MilestoneSizePt" => "Milestone size",
        "StandardOutlinePt" => "Outline width",
        "DelineatorLinePt" => "Delineator width",
        "GridLinePt" => "Grid line width",
        "MajorBoundaryPt" => "Major boundary width",
        "HatchPitchPt" => "Hatch spacing",
        "HatchLinePt" => "Hatch line width",
        _ => key,
    };

    private void CreateCustomStylesTab()
    {
#pragma warning disable CA1303 // Literal strings are UI labels in a WinForms dialog; localization is out of scope for R5.6a.
        TabPage tab = new TabPage("Custom styles")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        tab.Controls.Add(panel);

        // Left: the list of user-created styles. The list is the master; the
        // property panel on the right edits the selected object.
        _customStyleListGroup.Location = new Point(8, 8);
        _customStyleListGroup.Size = new Size(210, 300);
        _customStyleListLabel.Location = new Point(16, 24);
        _customStyleList.Location = new Point(16, 48);
        _customStyleList.Size = new Size(178, 240);

        _customStyleList.SelectedIndexChanged += CustomStyleList_SelectedIndexChanged;
        _customStyleListGroup.Controls.Add(_customStyleListLabel);
        _customStyleListGroup.Controls.Add(_customStyleList);
        panel.Controls.Add(_customStyleListGroup);

        // Empty-state hint (full width so the instruction stays readable).
        _customStyleListHint.AutoSize = true;
        _customStyleListHint.Location = new Point(8, 316);
        panel.Controls.Add(_customStyleListHint);

        // New-style entry is always available: it is not gated behind a
        // selection, so the first custom style can always be created.
        _customNewButton.Click += CustomNewButton_Click;
        _customNewButton.Location = new Point(16, 344);
        panel.Controls.Add(_customNewButton);
        _customNewKeyText.Location = new Point(112, 342);
        panel.Controls.Add(_customNewKeyText);

        // Right: the property panel for the selected style, grouped by concern
        // (Fill / Outline / Label / Size) rather than by table column.
        _customStyleEditGroup.Location = new Point(230, 8);
        _customStyleEditGroup.Size = new Size(470, 470);

        Label keyLbl = new Label
        {
            Text = _customStyleKeyLabel,
            Location = new Point(16, 26),
            Size = new Size(70, 20),
        };
        _customStyleKeyText.Location = new Point(92, 22);
        _customStyleKeyText.ReadOnly = true;

        GroupBox fillGroup = new GroupBox
        {
            Text = "Fill",
            Location = new Point(12, 52),
            Size = new Size(446, 78),
        };
        AddCustomColourRow(fillGroup, "Colour", _customStyleFillText, _customFillColorButton, _customFillSwatch, 22);
        Label patternLbl = new Label
        {
            Text = "Pattern",
            Location = new Point(16, 52),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        fillGroup.Controls.Add(patternLbl);
        _customHatchPatternCombo.Location = new Point(92, 48);
        fillGroup.Controls.Add(_customHatchPatternCombo);
        _customStyleEditGroup.Controls.Add(fillGroup);

        GroupBox outlineGroup = new GroupBox
        {
            Text = "Outline",
            Location = new Point(12, 138),
            Size = new Size(446, 106),
        };
        AddCustomColourRow(outlineGroup, "Colour", _customStyleStrokeText, _customStrokeColorButton, _customStrokeSwatch, 22);
        Label widthLbl = new Label
        {
            Text = "Width",
            Location = new Point(16, 52),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        outlineGroup.Controls.Add(widthLbl);
        _customOutlineWidthNumeric.Location = new Point(92, 48);
        outlineGroup.Controls.Add(_customOutlineWidthNumeric);
        Label dashLbl = new Label
        {
            Text = "Line style",
            Location = new Point(16, 80),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        outlineGroup.Controls.Add(dashLbl);
        _customStrokeDashCombo.Location = new Point(92, 76);
        outlineGroup.Controls.Add(_customStrokeDashCombo);
        _customStyleEditGroup.Controls.Add(outlineGroup);

        GroupBox labelGroup = new GroupBox
        {
            Text = "Label",
            Location = new Point(12, 252),
            Size = new Size(446, 118),
        };
        AddCustomColourRow(labelGroup, "Colour", _customStyleTextText, _customTextColorButton, _customTextSwatch, 22);
        Label defaultLbl = new Label
        {
            Text = "Default",
            Location = new Point(16, 52),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        labelGroup.Controls.Add(defaultLbl);
        _customDefaultLabelCombo.Location = new Point(92, 48);
        labelGroup.Controls.Add(_customDefaultLabelCombo);
        Label allowedLbl = new Label
        {
            Text = "Allowed",
            Location = new Point(16, 80),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        labelGroup.Controls.Add(allowedLbl);
        _customAllowedLabelList.Location = new Point(92, 74);
        labelGroup.Controls.Add(_customAllowedLabelList);
        _customStyleEditGroup.Controls.Add(labelGroup);

        GroupBox sizeGroup = new GroupBox
        {
            Text = "Size",
            Location = new Point(12, 378),
            Size = new Size(446, 78),
        };
        Label barLbl = new Label
        {
            Text = "Bar height",
            Location = new Point(16, 22),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        sizeGroup.Controls.Add(barLbl);
        _customActivityHeightNumeric.Location = new Point(92, 18);
        sizeGroup.Controls.Add(_customActivityHeightNumeric);
        Label milestoneLbl = new Label
        {
            Text = "Milestone",
            Location = new Point(16, 52),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        sizeGroup.Controls.Add(milestoneLbl);
        _customMilestoneSizeNumeric.Location = new Point(92, 48);
        sizeGroup.Controls.Add(_customMilestoneSizeNumeric);
        _customStyleEditGroup.Controls.Add(sizeGroup);

        // Delete only applies to a selected style; it lives inside the edit group.
        _customDeleteButton.Click += CustomDeleteButton_Click;
        _customDeleteButton.Location = new Point(16, 466);
        _customStyleEditGroup.Controls.Add(_customDeleteButton);

        _customStyleEditGroup.Controls.Add(keyLbl);
        _customStyleEditGroup.Controls.Add(_customStyleKeyText);

        PopulateCustomStyleVocabularies();
        panel.Controls.Add(_customStyleEditGroup);

        UpdateCustomStyleListHint();
#pragma warning restore CA1303
    }

    /// <summary>
    /// Adds a labelled colour row (hex text box, picker button, swatch) to a
    /// property group.
    /// </summary>
    /// <param name="parent">The group to add the row to.</param>
    /// <param name="label">The property label.</param>
    /// <param name="textBox">The hex text box.</param>
    /// <param name="pickerButton">The colour picker button.</param>
    /// <param name="swatch">The colour swatch.</param>
    /// <param name="top">The row's top offset inside the group.</param>
    private void AddCustomColourRow(
        Control parent,
        string label,
        TextBox textBox,
        Button pickerButton,
        Panel swatch,
        int top)
    {
        Label lbl = new Label
        {
            Text = label,
            Location = new Point(16, top + 3),
            Size = new Size(70, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        parent.Controls.Add(lbl);

        textBox.Location = new Point(92, top);
        textBox.TextChanged += CustomStyleColourTextChanged;

        pickerButton.Location = new Point(200, top - 1);
        pickerButton.Click += CustomFillColorButton_Click;

        swatch.Location = new Point(232, top);

        parent.Controls.Add(textBox);
        parent.Controls.Add(pickerButton);
        parent.Controls.Add(swatch);
    }

    /// <summary>
    /// Populates the closed vocabularies the custom-style editor offers: the
    /// fill patterns and the label positions. They are code-owned, so they are
    /// listed once rather than edited.
    /// </summary>
    private void PopulateCustomStyleVocabularies()
    {
        foreach (GanttHatchPattern pattern in Enum.GetValues<GanttHatchPattern>())
        {
            _ = _customHatchPatternCombo.Items.Add(new HatchPatternItem(pattern, DescribeHatchPattern(pattern)));
        }

        foreach (GanttLabelPosition position in Enum.GetValues<GanttLabelPosition>())
        {
            _ = _customDefaultLabelCombo.Items.Add(new LabelPositionItem(position, DescribeLabelPosition(position)));
            _ = _customAllowedLabelList.Items.Add(new LabelPositionItem(position, DescribeLabelPosition(position)));
        }

        foreach (GanttStrokeDash dash in Enum.GetValues<GanttStrokeDash>())
        {
            _ = _customStrokeDashCombo.Items.Add(new StrokeDashItem(dash, DescribeStrokeDash(dash)));
        }

        _customHatchPatternCombo.SelectedIndexChanged += CustomStylePropertyChanged;
        _customDefaultLabelCombo.SelectedIndexChanged += CustomStylePropertyChanged;
        _customStrokeDashCombo.SelectedIndexChanged += CustomStylePropertyChanged;
        _customOutlineWidthNumeric.ValueChanged += CustomStylePropertyChanged;
        _customActivityHeightNumeric.ValueChanged += CustomStylePropertyChanged;
        _customMilestoneSizeNumeric.ValueChanged += CustomStylePropertyChanged;
        _customAllowedLabelList.ItemCheck += CustomStyleAllowedLabelItemCheck;

        // The size controls are bounded by the same catalogue tokens the
        // global metric controls use, so a custom style cannot carry a size
        // the entity guide forbids and the control can refuse the value
        // rather than the writer.
        ApplyMetricRange(_customOutlineWidthNumeric, "StandardOutlinePt");
        ApplyMetricRange(_customActivityHeightNumeric, "ActivityHeightPt");
        ApplyMetricRange(_customMilestoneSizeNumeric, "MilestoneSizePt");
    }

    /// <summary>
    /// Bounds a numeric control by its catalogue metric token's range. The
    /// minimum is zero rather than the token's minimum, because zero is the
    /// style contract's "not applicable to this entity kind" value: a style
    /// used by an activity carries a zero milestone size and vice versa.
    /// </summary>
    /// <param name="numeric">The control to bound.</param>
    /// <param name="tokenName">The catalogue metric token name.</param>
    private static void ApplyMetricRange(NumericUpDown numeric, string tokenName)
    {
        GanttMetricToken? token = GanttCatalogues.Metrics
            .FirstOrDefault(candidate => string.Equals(candidate.Name, tokenName, StringComparison.Ordinal));
        if (token is null)
        {
            return;
        }

        numeric.Minimum = 0m;
        numeric.Maximum = (decimal)token.Maximum;
    }

    /// <summary>
    /// A hatch pattern paired with its user-facing label.
    /// </summary>
    /// <param name="Pattern">The pattern value.</param>
    /// <param name="Label">The display name.</param>
    private sealed record HatchPatternItem(GanttHatchPattern Pattern, string Label)
    {
        /// <inheritdoc />
        public override string ToString() => Label;
    }

    /// <summary>
    /// A label position paired with its user-facing label.
    /// </summary>
    /// <param name="Position">The position value.</param>
    /// <param name="Label">The display name.</param>
    private sealed record LabelPositionItem(GanttLabelPosition Position, string Label)
    {
        /// <inheritdoc />
        public override string ToString() => Label;
    }

    /// <summary>
    /// A stroke dash pattern paired with its user-facing label.
    /// </summary>
    /// <param name="Dash">The dash value.</param>
    /// <param name="Label">The display name.</param>
    private sealed record StrokeDashItem(GanttStrokeDash Dash, string Label)
    {
        /// <inheritdoc />
        public override string ToString() => Label;
    }

    /// <summary>
    /// Returns the user-facing name of a stroke dash pattern.
    /// </summary>
    /// <param name="dash">The dash pattern.</param>
    /// <returns>The display name.</returns>
    private static string DescribeStrokeDash(GanttStrokeDash dash) => dash switch
    {
        GanttStrokeDash.Solid => "Solid",
        GanttStrokeDash.SquareDot => "Square dot",
        GanttStrokeDash.RoundDot => "Round dot",
        GanttStrokeDash.Dash => "Dash",
        GanttStrokeDash.DashDot => "Dash dot",
        GanttStrokeDash.LongDash => "Long dash",
        _ => dash.ToString(),
    };

    /// <summary>
    /// Returns the user-facing name of a fill pattern.
    /// </summary>
    /// <param name="pattern">The pattern.</param>
    /// <returns>The display name.</returns>
    private static string DescribeHatchPattern(GanttHatchPattern pattern) => pattern switch
    {
        GanttHatchPattern.None => "Solid",
        GanttHatchPattern.ForwardDiagonal => "Forward diagonal",
        GanttHatchPattern.BackwardDiagonal => "Backward diagonal",
        GanttHatchPattern.Cross => "Cross",
        _ => pattern.ToString(),
    };

    /// <summary>
    /// Returns the user-facing name of a label position.
    /// </summary>
    /// <param name="position">The position.</param>
    /// <returns>The display name.</returns>
    private static string DescribeLabelPosition(GanttLabelPosition position) => position switch
    {
        GanttLabelPosition.None => "None",
        GanttLabelPosition.Auto => "Automatic",
        GanttLabelPosition.Left => "Left",
        GanttLabelPosition.Right => "Right",
        GanttLabelPosition.Inside => "Inside",
        GanttLabelPosition.TopLeft => "Top left",
        GanttLabelPosition.TopRight => "Top right",
        GanttLabelPosition.BottomLeft => "Bottom left",
        GanttLabelPosition.BottomRight => "Bottom right",
        GanttLabelPosition.DataPanelLeft => "Data panel",
        GanttLabelPosition.PlotCentre => "Plot centre",
        GanttLabelPosition.Both => "Both",
        _ => position.ToString(),
    };

    private void CustomStyleList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        _selectedCustomStyleKey = _customStyleList.SelectedItem as string;
        bool hasSelection = _selectedCustomStyleKey is not null;
        _customDeleteButton.Enabled = hasSelection;
        _customStyleEditGroup.Enabled = hasSelection;

        if (hasSelection && _selectedCustomStyleKey is { } key &&
            _customStyleEdits.TryGetValue(key, out CustomStyleState? state))
        {
            LoadCustomStyleIntoEditor(state);
        }
        else
        {
            ClearCustomStyleEditor();
        }

        UpdateCustomStyleListHint();
    }

    /// <summary>
    /// Loads a style's full property state into the editor controls. The
    /// controls are written directly rather than through their change events,
    /// so populating the editor does not itself mark the style as edited.
    /// </summary>
    /// <param name="state">The style state to load.</param>
    private void LoadCustomStyleIntoEditor(CustomStyleState state)
    {
        _suppressCustomStyleSync = true;
        try
        {
            LoadCustomStyleIntoEditorCore(state);
        }
        finally
        {
            _suppressCustomStyleSync = false;
        }
    }

    /// <summary>
    /// Writes a style's property state into the editor controls.
    /// </summary>
    /// <param name="state">The style state to load.</param>
    private void LoadCustomStyleIntoEditorCore(CustomStyleState state)
    {
        _customStyleKeyText.Text = _selectedCustomStyleKey ?? string.Empty;
        _customStyleFillText.Text = state.Fill;
        _customStyleStrokeText.Text = state.Stroke;
        _customStyleTextText.Text = state.Text;
        _customFillSwatch.BackColor = ParseColourOrDefault(state.Fill);
        _customStrokeSwatch.BackColor = ParseColourOrDefault(state.Stroke);
        _customTextSwatch.BackColor = ParseColourOrDefault(state.Text);

        _customHatchPatternCombo.SelectedItem =
            _customHatchPatternCombo.Items.Cast<HatchPatternItem>()
                .FirstOrDefault(item => item.Pattern == state.HatchPattern);
        _customStrokeDashCombo.SelectedItem =
            _customStrokeDashCombo.Items.Cast<StrokeDashItem>()
                .FirstOrDefault(item => item.Dash == state.StrokeDash);
        _customDefaultLabelCombo.SelectedItem =
            _customDefaultLabelCombo.Items.Cast<LabelPositionItem>()
                .FirstOrDefault(item => item.Position == state.DefaultLabelPosition);

        for (int index = 0; index < _customAllowedLabelList.Items.Count; index++)
        {
            bool isChecked = _customAllowedLabelList.Items[index] is LabelPositionItem item
                && state.AllowedLabelPositions.Contains(item.Position);
            _customAllowedLabelList.SetItemChecked(index, isChecked);
        }

        _customOutlineWidthNumeric.Value = ClampToNumericRange(state.StandardOutlinePt, _customOutlineWidthNumeric);
        _customActivityHeightNumeric.Value = ClampToNumericRange(state.ActivityHeightPt, _customActivityHeightNumeric);
        _customMilestoneSizeNumeric.Value = ClampToNumericRange(state.MilestoneSizePt, _customMilestoneSizeNumeric);
    }

    /// <summary>
    /// Clamps a stored value into a numeric control's range, so a stored
    /// value outside the control's bounds cannot throw on assignment.
    /// </summary>
    /// <param name="value">The stored value.</param>
    /// <param name="numeric">The target control.</param>
    /// <returns>The clamped value.</returns>
    private static decimal ClampToNumericRange(double value, NumericUpDown numeric)
    {
        if (!double.IsFinite(value))
        {
            return numeric.Minimum;
        }

        decimal candidate = (decimal)Math.Clamp(value, (double)numeric.Minimum, (double)numeric.Maximum);
        return candidate;
    }

    private void ClearCustomStyleEditor()
    {
        _suppressCustomStyleSync = true;
        try
        {
            ClearCustomStyleEditorCore();
        }
        finally
        {
            _suppressCustomStyleSync = false;
        }
    }

    /// <summary>
    /// Resets the editor controls to their empty state.
    /// </summary>
    private void ClearCustomStyleEditorCore()
    {
        _selectedCustomStyleKey = null;
        _customDeleteButton.Enabled = false;
        _customStyleEditGroup.Enabled = false;
        _customStyleKeyText.Text = string.Empty;
        _customStyleFillText.Text = string.Empty;
        _customStyleStrokeText.Text = string.Empty;
        _customStyleTextText.Text = string.Empty;
        _customFillSwatch.BackColor = SystemColors.Control;
        _customStrokeSwatch.BackColor = SystemColors.Control;
        _customTextSwatch.BackColor = SystemColors.Control;
        _customHatchPatternCombo.SelectedIndex = -1;
        _customDefaultLabelCombo.SelectedIndex = -1;
        _customStrokeDashCombo.SelectedIndex = -1;
        for (int index = 0; index < _customAllowedLabelList.Items.Count; index++)
        {
            _customAllowedLabelList.SetItemChecked(index, false);
        }

        _customOutlineWidthNumeric.Value = 0m;
        _customActivityHeightNumeric.Value = 0m;
        _customMilestoneSizeNumeric.Value = 0m;
    }

#pragma warning disable CA1303 // Literal strings are UI labels in a WinForms dialog; localization is out of scope for R5.6a.
    private void CustomNewButton_Click(object? sender, EventArgs e)
    {
        string newKey = _customNewKeyText.Text.Trim().ToUpperInvariant();
        _customNewKeyText.Text = string.Empty;

        if (string.IsNullOrWhiteSpace(newKey))
        {
            _warningLabel.Text = "Enter a key for the new custom style.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _ = _customNewKeyText.Focus();
            return;
        }

        if (IsBuiltInCustomKey(newKey))
        {
            _warningLabel.Text = $"Style key '{newKey}' is a built-in preset and cannot be used as a custom style.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            return;
        }

        if (_customStyleEdits.ContainsKey(newKey))
        {
            _warningLabel.Text = $"Custom style '{newKey}' already exists.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            return;
        }

        _warningLabel.Visible = false;

        // The new style starts from the As-Planned tokens so it renders as a
        // visible entity rather than a degenerate one, and it permits the
        // label positions a span or a milestone both accept.
        _customStyleEdits[newKey] = new CustomStyleState(
            "#FFFFFF",
            "#000000",
            "#000000",
            GanttHatchPattern.None,
            GanttCatalogues.MetricDefault("StandardOutlinePt"),
            GanttCatalogues.MetricDefault("ActivityHeightPt"),
            GanttCatalogues.MetricDefault("MilestoneSizePt"),
            GanttLabelPosition.Auto,
            new HashSet<GanttLabelPosition>
            {
                GanttLabelPosition.Auto,
                GanttLabelPosition.Left,
                GanttLabelPosition.Right,
                GanttLabelPosition.Inside,
                GanttLabelPosition.None,
            },
            GanttStrokeDash.Solid);
        _ = _customStyleList.Items.Add(newKey);
        _customStyleList.SelectedItem = newKey;
    }
#pragma warning restore CA1303

    /// <summary>
    /// Determines whether two permitted-label sets are the same. A null set and
    /// an empty set are the same (neither permits a position); the comparison
    /// is ordinal so the two sets cannot disagree over case.
    /// </summary>
    /// <param name="left">The first set.</param>
    /// <param name="right">The second set.</param>
    /// <returns>Whether the sets permit the same positions.</returns>
    private static bool SameLabelPositions(
        IReadOnlySet<GanttLabelPosition>? left,
        IReadOnlySet<GanttLabelPosition>? right)
    {
        int leftCount = left?.Count ?? 0;
        int rightCount = right?.Count ?? 0;
        if (leftCount != rightCount)
        {
            return false;
        }

        if (leftCount == 0)
        {
            return true;
        }

        foreach (GanttLabelPosition position in left!)
        {
            if (!right!.Contains(position))
            {
                return false;
            }
        }

        return true;
    }

    private static bool IsBuiltInCustomKey(string key) =>
        GanttCatalogues.StylePresets.Any(p =>
            string.Equals(p.StyleKey, key, StringComparison.Ordinal));

    private void CustomDeleteButton_Click(object? sender, EventArgs e)
    {
        if (_selectedCustomStyleKey is not { } key)
        {
            return;
        }

        DialogResult confirm = MessageBox.Show(
            $"Delete custom style '{key}'?",
            "Delete custom style",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        if (confirm != DialogResult.Yes)
        {
            return;
        }

        _ = _customStyleEdits.Remove(key);
        _customStyleList.Items.Remove(key);
        ClearCustomStyleEditor();
        UpdateCustomStyleListHint();
    }

    private void UpdateCustomStyleListHint()
    {
        bool empty = _customStyleList.Items.Count == 0;
        _customStyleListHint.Visible = empty;
        _customStyleList.Visible = !empty;
        _customStyleListLabel.Visible = !empty;
    }

    private void CustomFillColorButton_Click(object? sender, EventArgs e) => PickCustomColour(_customStyleFillText, _customFillSwatch);
    private void CustomStrokeColorButton_Click(object? sender, EventArgs e) => PickCustomColour(_customStyleStrokeText, _customStrokeSwatch);
    private void CustomTextColorButton_Click(object? sender, EventArgs e) => PickCustomColour(_customStyleTextText, _customTextSwatch);

    private void PickCustomColour(TextBox textBox, Panel swatch)
    {
        _colorDialog.Color = ColourHex.TryParse(textBox.Text, out ColourHex? initial)
            ? Color.FromArgb((int)initial!.ARGB)
            : SystemColors.Control;

        if (_colorDialog.ShowDialog(this) == DialogResult.OK)
        {
            Color color = _colorDialog.Color;
            string hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            textBox.Text = hex;
            swatch.BackColor = color;
        }
    }

    private void CustomStyleColourTextChanged(object? sender, EventArgs e)
    {
        if (_selectedCustomStyleKey is null || _suppressCustomStyleSync)
        {
            return;
        }

        SyncSelectedCustomStyleState();
    }

    /// <summary>
    /// Handles a change to one of the custom style's non-colour properties
    /// (the fill pattern, the outline width, the sizes, the default label
    /// position) by writing the whole style state back to the edit model.
    /// </summary>
    /// <param name="sender">The changed control.</param>
    /// <param name="e">The event arguments.</param>
    private void CustomStylePropertyChanged(object? sender, EventArgs e)
    {
        if (_selectedCustomStyleKey is null || _suppressCustomStyleSync)
        {
            return;
        }

        SyncSelectedCustomStyleState();
    }

    /// <summary>
    /// Handles a check/uncheck in the permitted-label list. The list's
    /// <c>ItemCheck</c> event fires <em>before</em> the check state changes, so
    /// the new set is computed from the event's own index and new value rather
    /// than read back from the control.
    /// </summary>
    /// <param name="sender">The list.</param>
    /// <param name="e">The event arguments.</param>
    private void CustomStyleAllowedLabelItemCheck(object? sender, ItemCheckEventArgs e)
    {
        if (_selectedCustomStyleKey is null || _suppressCustomStyleSync)
        {
            return;
        }

        HashSet<GanttLabelPosition> selected = [];
        for (int index = 0; index < _customAllowedLabelList.Items.Count; index++)
        {
            bool isChecked = index == e.Index
                ? e.NewValue == CheckState.Checked
                : _customAllowedLabelList.GetItemChecked(index);
            if (isChecked && _customAllowedLabelList.Items[index] is LabelPositionItem item)
            {
                _ = selected.Add(item.Position);
            }
        }

        _customStyleEdits[_selectedCustomStyleKey] = new CustomStyleState(
            _customStyleFillText.Text,
            _customStyleStrokeText.Text,
            _customStyleTextText.Text,
            SelectedHatchPattern(),
            (double)_customOutlineWidthNumeric.Value,
            (double)_customActivityHeightNumeric.Value,
            (double)_customMilestoneSizeNumeric.Value,
            SelectedDefaultLabelPosition(),
            selected,
            SelectedStrokeDash());
    }

    /// <summary>
    /// Copies the current editor values into the edit model for the selected
    /// style. The whole object is written, so every property the editor owns
    /// is carried in the edit and nothing is left to a stale default.
    /// </summary>
    private void SyncSelectedCustomStyleState()
    {
        if (_selectedCustomStyleKey is null)
        {
            return;
        }

        _customStyleEdits[_selectedCustomStyleKey] = new CustomStyleState(
            _customStyleFillText.Text,
            _customStyleStrokeText.Text,
            _customStyleTextText.Text,
            SelectedHatchPattern(),
            (double)_customOutlineWidthNumeric.Value,
            (double)_customActivityHeightNumeric.Value,
            (double)_customMilestoneSizeNumeric.Value,
            SelectedDefaultLabelPosition(),
            SelectedAllowedLabelPositions(),
            SelectedStrokeDash());
    }

    /// <summary>
    /// Returns the fill pattern currently selected in the editor.
    /// </summary>
    /// <returns>The selected pattern, or None when nothing is selected.</returns>
    private GanttHatchPattern SelectedHatchPattern() =>
        _customHatchPatternCombo.SelectedItem is HatchPatternItem item
            ? item.Pattern
            : GanttHatchPattern.None;

    /// <summary>
    /// Returns the default label position currently selected in the editor.
    /// </summary>
    /// <returns>The selected position, or Auto when nothing is selected.</returns>
    private GanttLabelPosition SelectedDefaultLabelPosition() =>
        _customDefaultLabelCombo.SelectedItem is LabelPositionItem item
            ? item.Position
            : GanttLabelPosition.Auto;

    /// <summary>
    /// Returns the stroke dash pattern currently selected in the editor.
    /// </summary>
    /// <returns>The selected dash, or Solid when nothing is selected.</returns>
    private GanttStrokeDash SelectedStrokeDash() =>
        _customStrokeDashCombo.SelectedItem is StrokeDashItem item
            ? item.Dash
            : GanttStrokeDash.Solid;

    /// <summary>
    /// Returns the checked label positions in the editor.
    /// </summary>
    /// <returns>The checked positions.</returns>
    private HashSet<GanttLabelPosition> SelectedAllowedLabelPositions()
    {
        HashSet<GanttLabelPosition> selected = [];
        for (int index = 0; index < _customAllowedLabelList.Items.Count; index++)
        {
            if (_customAllowedLabelList.GetItemChecked(index)
                && _customAllowedLabelList.Items[index] is LabelPositionItem item)
            {
                _ = selected.Add(item.Position);
            }
        }

        return selected;
    }

    private void PopulateFromRequest(ThemeEditRequest request)
    {
        // Colours. The master–detail view shows one style at a time, so the
        // workbook's current colours for EVERY style are held in the pending
        // map; the detail panel reads from it when a style is selected.
        //
        // Every built-in style is seeded from the code-owned preset FIRST, so
        // a style the request does not mention keeps its code-owned colours
        // rather than losing them: the writer treats an absent edit as "use
        // the preset", so an unseeded style would silently drop a colour the
        // user had already chosen.
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            _entityStyleColours[preset.StyleKey] = new EntityStyleColours(
                preset.FillColour,
                preset.StrokeColour,
                preset.TextColour);
        }

        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            _entityStyleColours[edit.StyleKey] = new EntityStyleColours(
                edit.FillColour,
                edit.StrokeColour,
                edit.TextColour);
        }

        if (_selectedEntityStyleKey is { } selectedKey
            && _entityStyleColours.TryGetValue(selectedKey, out EntityStyleColours? selected))
        {
            _entityFillText.Text = selected.Fill;
            _entityStrokeText.Text = selected.Stroke;
            _entityTextText.Text = selected.Text;
            UpdateSwatch("entityFill", ParseColourOrDefault(selected.Fill));
            UpdateSwatch("entityStroke", ParseColourOrDefault(selected.Stroke));
            UpdateSwatch("entityText", ParseColourOrDefault(selected.Text));
        }

        // Metrics
        foreach (KeyValuePair<string, double> kvp in request.MetricEdits)
        {
            if (_metricNumericUpDowns.TryGetValue(kvp.Key, out NumericUpDown? nud))
            {
                nud.Value = (decimal)kvp.Value;
            }
        }

        // The chart-chrome colours: the workbook's current values, so the
        // dialog shows what the chart actually renders rather than the
        // catalogue default.
        foreach (KeyValuePair<string, string> pair in request.ChromeColourEdits
            ?? new Dictionary<string, string>(StringComparer.Ordinal))
        {
            if (_chromeColourTextBoxes.TryGetValue(pair.Key, out TextBox? txt))
            {
                txt.Text = pair.Value;
                UpdateSwatch(pair.Key, ParseColourOrDefault(pair.Value));
            }
        }

        // Custom styles: load saved styles into the edit dictionary and list so
        // they are visible on the tab and are not spuriously deleted on OK.
        // The whole property set is loaded, not just the colours: an editor
        // that showed only the colours would overwrite the stored pattern,
        // sizes, and label vocabulary with defaults the user never chose.
        foreach (CustomStyleDef def in request.CustomStyleDefs ?? [])
        {
            _customStyleEdits[def.StyleKey] = new CustomStyleState(
                def.FillColour,
                def.StrokeColour,
                def.TextColour,
                def.HatchPattern,
                def.StandardOutlinePt,
                def.ActivityHeightPt,
                def.MilestoneSizePt,
                def.DefaultLabelPosition,
                def.AllowedLabelPositions,
                def.StrokeDash);
            _ = _customStyleList.Items.Add(def.StyleKey);
        }
        UpdateCustomStyleListHint();
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        // Validate all colour text boxes
        foreach (KeyValuePair<string, TextBox> kvp in _colorTextBoxes)
        {
            if (!string.IsNullOrEmpty(kvp.Value.Text) && !ColourHex.TryParse(kvp.Value.Text, out _))
            {
                _warningLabel.Text = $"Invalid colour '{kvp.Value.Text}' for {kvp.Key}. Use uppercase #RRGGBB.";
                _warningLabel.Visible = true;
                _tabControl.SelectedTab = _tabControl.TabPages[0];
                _ = kvp.Value.Focus();
                DialogResult = DialogResult.None;
                return;
            }
        }

        // Validate custom style colours
        if (!string.IsNullOrEmpty(_customStyleFillText.Text) && !ColourHex.TryParse(_customStyleFillText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style fill colour '" + _customStyleFillText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleFillText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleStrokeText.Text) && !ColourHex.TryParse(_customStyleStrokeText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style stroke colour '" + _customStyleStrokeText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleStrokeText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleTextText.Text) && !ColourHex.TryParse(_customStyleTextText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style text colour '" + _customStyleTextText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleTextText.Focus();
            DialogResult = DialogResult.None;
            return;
        }

        // Validate the chart-chrome colours. A malformed one is refused with
        // the offending row focused, so the user fixes the colour they typed
        // rather than finding a silent substitution in the rendered chart.
        foreach (KeyValuePair<string, TextBox> kvp in _chromeColourTextBoxes)
        {
            if (!string.IsNullOrEmpty(kvp.Value.Text) && !ColourHex.TryParse(kvp.Value.Text, out _))
            {
                _warningLabel.Text = "Invalid chart colour '" + kvp.Value.Text + "' for "
                    + GanttChartChrome.DisplayLabelFor(ChromeRoleFor(kvp.Key))
                    + ". Use uppercase #RRGGBB.";
                _warningLabel.Visible = true;
                _tabControl.SelectedTab = _tabControl.TabPages[1];
                _ = kvp.Value.Focus();
                DialogResult = DialogResult.None;
                return;
            }
        }

        Dictionary<string, double> metricEdits = [];
        foreach (KeyValuePair<string, NumericUpDown> kvp in _metricNumericUpDowns)
        {
            metricEdits[kvp.Key] = (double)kvp.Value.Value;
        }

        // Build request
        // The pending colours for the style on screen are captured first, so
        // an edit the user made just before pressing OK is not lost.
        CaptureSelectedEntityColours();

        List<ThemeStyleEdit> styleEdits = [];
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            if (!_entityStyleColours.TryGetValue(preset.StyleKey, out EntityStyleColours? colours))
            {
                continue;
            }

            string fill = colours.Fill;
            string stroke = colours.Stroke;
            string text = colours.Text;

            if (fill.Length > 0 || stroke.Length > 0 || text.Length > 0)
            {
                // The line type is carried only for the delineator, whose dash
                // the reader projects from the row rather than exact-matching
                // it against the preset. Every other built-in style keeps its
                // code-owned dash.
                GanttStrokeDash? dash = string.Equals(preset.StyleKey, "DefaultDelineator", StringComparison.Ordinal)
                    ? SelectedEntityDash()
                    : null;

                styleEdits.Add(new ThemeStyleEdit(preset.StyleKey, fill, stroke, text, dash));
            }
        }

        List<CustomStyleEdit> customStyleEdits = [];
        HashSet<string> originalKeys = new HashSet<string>(_originalCustomDefs.Select(d => d.StyleKey), StringComparer.Ordinal);

        // Deletes: original custom styles no longer in _customStyleEdits
        foreach (CustomStyleDef original in _originalCustomDefs)
        {
            if (!_customStyleEdits.ContainsKey(original.StyleKey))
            {
                customStyleEdits.Add(new CustomStyleEdit(CustomStyleOperation.Delete, original.StyleKey));
            }
        }

        // Creates and updates: current custom styles
        foreach (KeyValuePair<string, CustomStyleState> kvp in _customStyleEdits)
        {
            string key = kvp.Key;
            CustomStyleState state = kvp.Value;
            CustomStyleDef? original = _originalCustomDefs.FirstOrDefault(d => d.StyleKey == key);
            if (original is null)
            {
                customStyleEdits.Add(new CustomStyleEdit(
                    CustomStyleOperation.Create,
                    key,
                    DisplayName: key,
                    FillColour: state.Fill,
                    StrokeColour: state.Stroke,
                    TextColour: state.Text,
                    HatchPattern: state.HatchPattern,
                    StandardOutlinePt: state.StandardOutlinePt,
                    ActivityHeightPt: state.ActivityHeightPt,
                    MilestoneSizePt: state.MilestoneSizePt,
                    DefaultLabelPosition: state.DefaultLabelPosition,
                    AllowedLabelPositions: state.AllowedLabelPositions,
                    StrokeDash: state.StrokeDash));
            }
            else if (original.FillColour != state.Fill
                || original.StrokeColour != state.Stroke
                || original.TextColour != state.Text
                || original.HatchPattern != state.HatchPattern
                || original.StandardOutlinePt != state.StandardOutlinePt
                || original.ActivityHeightPt != state.ActivityHeightPt
                || original.MilestoneSizePt != state.MilestoneSizePt
                || original.DefaultLabelPosition != state.DefaultLabelPosition
                || original.StrokeDash != state.StrokeDash
                || (original.AllowedLabelPositions is not null
                    && !SameLabelPositions(original.AllowedLabelPositions, state.AllowedLabelPositions)))
            {
                // A change on an existing key is an upsert by key: the key is
                // unchanged, so it can never collide with another key
                // (R5.6a D1). A Rename with identical old/new key would
                // falsely fail the request's duplicate-key check.
                customStyleEdits.Add(new CustomStyleEdit(
                    CustomStyleOperation.Update,
                    key,
                    DisplayName: key,
                    FillColour: state.Fill,
                    StrokeColour: state.Stroke,
                    TextColour: state.Text,
                    HatchPattern: state.HatchPattern,
                    StandardOutlinePt: state.StandardOutlinePt,
                    ActivityHeightPt: state.ActivityHeightPt,
                    MilestoneSizePt: state.MilestoneSizePt,
                    DefaultLabelPosition: state.DefaultLabelPosition,
                    AllowedLabelPositions: state.AllowedLabelPositions,
                    StrokeDash: state.StrokeDash));
            }
        }

        // The chart-chrome colours the user set. A colour left at the
        // catalogue default is still carried: the setting's stored value is
        // the workbook's own, and the writer preserves it either way.
        Dictionary<string, string> chromeColourEdits = [];
        foreach (KeyValuePair<string, TextBox> kvp in _chromeColourTextBoxes)
        {
            string text = kvp.Value.Text.Trim();
            if (text.Length > 0)
            {
                chromeColourEdits[kvp.Key] = text;
            }
        }

        ThemeEditRequest request = new ThemeEditRequest(
            styleEdits,
            metricEdits,
            customStyleEdits,
            null,
            chromeColourEdits);

        if (!ThemeEditRequest.TryValidate(request, out _, out string? msg))
        {
            _warningLabel.Text = msg ?? "Invalid theme edit.";
            _warningLabel.Visible = true;
            DialogResult = DialogResult.None;
            return;
        }

        IReadOnlyList<string> warnings = ThemeEditRequest.ContrastWarnings(request);
        if (warnings.Count > 0)
        {
            _warningLabel.Text = string.Join("; ", warnings);
            _warningLabel.Visible = true;
        }

        _validatedRequest = request;
        DialogResult = DialogResult.OK;
    }

    /// <summary>
    /// Resets every editable theme field to the built-in default so that
    /// clicking OK overwrites all custom theme entries with the code-owned
    /// presets and metric defaults. Runs on the STA thread; never throws.
    /// </summary>
    private void ResetButton_Click(object? sender, EventArgs e)
    {
        // Resetting overwrites every custom theme entry, so confirm first.
        DialogResult confirm = MessageBox.Show(
            "Reset all theme colours and metric defaults to the built-in values? This overwrites all custom entries.",
            "Reset theme to default",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        if (confirm != DialogResult.Yes)
        {
            return;
        }

        // Empty colour boxes mean "use the preset default".
        foreach (TextBox txt in _colorTextBoxes.Values)
        {
            txt.Text = string.Empty;
        }

        // Reset every editable metric to its catalogue default.
        foreach (KeyValuePair<string, NumericUpDown> kvp in _metricNumericUpDowns)
        {
            GanttMetricToken? token = GanttCatalogues.Metrics
                .FirstOrDefault(t => string.Equals(t.Name, kvp.Key, StringComparison.Ordinal));
            if (token is not null)
            {
                kvp.Value.Value = (decimal)token.DefaultValue;
            }
        }

        // Refresh swatches to the unset/default look.
        Color defaultSwatch = SystemColors.Control;
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            if (preset.ColourCapability.HasFlag(EntityColourCapability.Fill) ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                UpdateSwatch($"fill_{preset.StyleKey}", defaultSwatch);
            }
            if (preset.ColourCapability.HasFlag(EntityColourCapability.Stroke) ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                UpdateSwatch($"stroke_{preset.StyleKey}", defaultSwatch);
            }
            if (preset.TextColour.Length > 0 ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                UpdateSwatch($"text_{preset.StyleKey}", defaultSwatch);
            }
        }

        _warningLabel.Visible = false;
    }
}
