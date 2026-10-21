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
    private readonly Label _entityLineWidthLabel = new();

    /// <summary>
    /// The pending colour values for each built-in style the user has edited
    /// or visited, keyed by the style key. The master-detail view shows one
    /// style at a time, so the values for the styles that are not on screen
    /// live here rather than in a text box.
    /// </summary>
    private readonly Dictionary<string, EntityStyleColours> _entityStyleColours = [];

    /// <summary>
    /// The pending line type for each built-in style, keyed by the style key.
    /// <para>
    /// Only the delineator's line type is user-editable, but the value lives
    /// in a map rather than in the combo alone for the same reason the
    /// colours do: the master-detail view shows one style at a time, so a
    /// value the user chose must survive a move to another style and back.
    /// It also carries the workbook's current value in, so the dialog opens
    /// on the line type the chart actually draws rather than the code-owned
    /// default.
    /// </para>
    /// </summary>
    private readonly Dictionary<string, GanttStrokeDash> _entityStyleDashes = [];

    /// <summary>
    /// The style key whose properties the detail panel is currently showing.
    /// </summary>
    private string? _selectedEntityStyleKey;

    private const string _customStyleKeyLabel = "Style key:";
    private const string _customStyleEmptyHint = "No custom styles defined yet. Enter a key and click New custom style.";

    /// <summary>
    /// The width of the label column in every property group. The groups are
    /// as wide as their container, so the label column is what keeps the
    /// labels of one style aligned with the labels of the next.
    /// </summary>
    private const int _propertyLabelColumnWidth = 130;

    /// <summary>
    /// The vertical gap between the property groups in a detail panel.
    /// </summary>
    private const int _propertyGroupGap = 10;

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
        IntegralHeight = true,
        Size = new Size(140, 96),
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
            _entityLineWidthLabel.Dispose();
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

        // The client area is sized to the CONTENT, not to a guess. The
        // master-detail view needs a master column (250pt) plus a detail
        // panel that holds a 130pt label column and a 160pt value cluster
        // (colour text box, picker button and swatch) — about 330pt — plus
        // the tab and form padding. 700pt is that sum with a little room,
        // so no tab needs a horizontal scroll bar and none of the property
        // values are stranded in a wide empty strip.
        ClientSize = new Size(700, 640);
        MinimumSize = new Size(700, 640);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = SystemFonts.MessageBoxFont;

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

        // The warning label uses the dialog's font, so it reads at the same
        // size as everything else rather than as a smaller footnote.
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

        // This is the ONLY title the detail panel carries. It opens on the
        // empty-state hint in the muted colour and turns into the selected
        // style's name once a row is chosen; an earlier revision added a second,
        // grey hint label on top of it, so the empty state showed the hint twice
        // and kept showing it after a style was selected.
        _entityStyleTitleLabel.Text = "Select a style to see its properties.";
        _entityStyleTitleLabel.ForeColor = SystemColors.GrayText;

        // The groups are stacked in a table rather than a flow panel: a flow
        // panel gives every child its own preferred width, so the groups were
        // 380px wide inside a 606px panel and left a wide empty strip on the
        // right. A table's single percent-width column gives every group the
        // panel's full width, and each group still hugs its own content.
        TableLayoutPanel body = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            AutoScroll = true,
        };
        _ = body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        for (int row = 0; row < 4; row++)
        {
            _ = body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        // The body is added BEFORE the title even though the title belongs at
        // the top of the panel: docked controls are laid out in REVERSE
        // z-order, so the last docked control added claims its strip first. In
        // the other order the Fill body consumed the entire panel and the Top
        // title was laid out over it at the same origin (both measured {8,8}),
        // leaving the panel's background to paint over the title text.
        _entityStyleDetailPanel.Controls.Add(body);
        _entityStyleDetailPanel.Controls.Add(_entityStyleTitleLabel);

        body.Controls.Add(BuildEntityFillGroup(), 0, 0);
        body.Controls.Add(BuildEntityOutlineGroup(), 0, 1);
        body.Controls.Add(BuildEntityLabelGroup(), 0, 2);
        body.Controls.Add(BuildEntitySizeGroup(), 0, 3);
    }

    /// <summary>
    /// The Fill group: the fill colour and the fill pattern.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntityFillGroup()
    {
        GroupBox group = NewPropertyGroup("Fill");
        TableLayoutPanel table = NewPropertyTable();
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

        AddClusterRow(table, "Colour", 0, _entityFillText, _entityFillPicker, _entityFillSwatch);

        _entityPatternLabel.AutoSize = true;
        _entityPatternLabel.ForeColor = SystemColors.GrayText;
        AddPropertyRow(table, "Pattern", _entityPatternLabel, 1);

        _colorTextBoxes["entityFill"] = _entityFillText;
        _colorPickerButtons["entityFill"] = _entityFillPicker;
        _colorSwatches["entityFill"] = _entityFillSwatch;

        return group;
    }

    /// <summary>
    /// Creates a property group: a titled box that spans its container's width
    /// and hugs the height of the properties it holds.
    /// </summary>
    /// <param name="caption">The group's caption.</param>
    /// <returns>The new group.</returns>
    /// <remarks>
    /// The group is docked <c>Top</c> rather than left to its own preferred
    /// size, so it fills the width of the detail panel and every group in the
    /// panel is the same width. Its height still comes from its content, so
    /// the box hugs the rows it holds.
    /// </remarks>
    private static GroupBox NewPropertyGroup(string caption) => new()
    {
        Text = caption,
        AutoSize = true,
        Dock = DockStyle.Top,
        Margin = new Padding(0, 0, 0, _propertyGroupGap),
    };

    /// <summary>
    /// Creates the two-column table every property group uses: a fixed label
    /// column and a value column that takes the rest of the width.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The table DOES fill its group box. A group is measured by the table's
    /// preferred size, and a table docked with <c>Dock.Top</c> measures empty —
    /// the group then collapses to its border alone (a 6px sliver) and every
    /// property in it becomes unreadable. Docked with <c>Dock.Fill</c> the
    /// table keeps reporting its content height, so the group grows around it
    /// while the table still spans the group's width.
    /// </para>
    /// <para>
    /// The table's <c>AutoSizeMode</c> must additionally be
    /// <c>GrowAndShrink</c>, not the default <c>GrowOnly</c>.
    /// A GrowOnly control "will not shrink smaller than the value specified by
    /// its Size", and a <see cref="TableLayoutPanel"/> starts at the default
    /// 200x100 — so with the default mode every table under 100px of content
    /// sat at exactly 100px with its LAST implicit row stretched to absorb the
    /// slack. GrowAndShrink ignores the current Size and hugs the content at
    /// any DPI.
    /// </para>
    /// </remarks>
    private static TableLayoutPanel NewPropertyTable()
    {
        TableLayoutPanel table = new()
        {
            ColumnCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
        };
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, _propertyLabelColumnWidth));
        _ = table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        return table;
    }

    /// <summary>
    /// Adds a labelled property row whose value is a single control.
    /// </summary>
    /// <param name="table">The group's property table.</param>
    /// <param name="label">The property label.</param>
    /// <param name="value">The value control.</param>
    /// <param name="row">The row the control sits in.</param>
    private static void AddPropertyRow(
        TableLayoutPanel table,
        string label,
        Control value,
        int row)
    {
        EnsureAutoSizeRow(table, row);

        table.Controls.Add(new Label
        {
            Text = label,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
        }, 0, row);

        // The value is anchored to the right of the value column, so every
        // property in the group lines up on the same right-hand edge — the
        // rhythm the Chart tab's metric rows already use.
        //
        // The value carries no margin: a taller control than the row's label
        // — a numeric up-down — would otherwise be laid out past the row's
        // height and clipped by the group box.
        value.Margin = new Padding(0);
        value.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        table.Controls.Add(value, 1, row);
    }

    /// <summary>
    /// Ensures the row at the given index sizes itself to its content.
    /// </summary>
    /// <param name="table">The group's property table.</param>
    /// <param name="row">The row index.</param>
    /// <remarks>
    /// A row that is not explicitly auto-sized keeps the height of the first
    /// control laid out in it, so a taller control in a later row — a numeric
    /// up-down beside a label — was clipped by the row above it. Every row is
    /// given an auto-size style so it grows to whatever it holds.
    /// </remarks>
    private static void EnsureAutoSizeRow(TableLayoutPanel table, int row)
    {
        while (table.RowStyles.Count <= row)
        {
            _ = table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        table.RowStyles[row] = new RowStyle(SizeType.AutoSize);
    }

    /// <summary>
    /// Adds a labelled property row whose value is a cluster of controls, such
    /// as a colour's text box, picker button and swatch.
    /// </summary>
    /// <param name="table">The group's property table.</param>
    /// <param name="label">The property label.</param>
    /// <param name="controls">The value controls, in display order.</param>
    /// <param name="row">The row the cluster sits in.</param>
    private static void AddClusterRow(
        TableLayoutPanel table,
        string label,
        int row,
        params Control[] controls)
    {
        EnsureAutoSizeRow(table, row);

        table.Controls.Add(new Label
        {
            Text = label,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
        }, 0, row);

        FlowLayoutPanel cluster = new()
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Margin = new Padding(0),
        };

        foreach (Control control in controls)
        {
            control.Margin = new Padding(0, 0, 4, 0);
            cluster.Controls.Add(control);
        }

        table.Controls.Add(cluster, 1, row);
    }
    /// <summary>
    /// The Outline group: the outline colour, its width, and its line type.
    /// </summary>
    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private GroupBox BuildEntityOutlineGroup()
    {
        GroupBox group = NewPropertyGroup("Outline");
        TableLayoutPanel table = NewPropertyTable();
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

        AddClusterRow(table, "Colour", 0, _entityStrokeText, _entityStrokePicker, _entityStrokeSwatch);

        _entityOutlineWidthLabel.AutoSize = true;
        _entityOutlineWidthLabel.ForeColor = SystemColors.GrayText;
        AddPropertyRow(table, "Width", _entityOutlineWidthLabel, 1);

        _entityDashLabel.AutoSize = true;
        _entityDashLabel.ForeColor = SystemColors.GrayText;
        _entityDashCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        _entityDashCombo.Width = 140;

        // The combo is shown for the delineator, whose line type the user may
        // change; the label carries the code-owned line type of every other
        // built-in style. Only one of the two is ever visible, so they share
        // a cluster: a hidden control takes no space in a flow panel, and the
        // visible one keeps the row's right-hand edge.
        AddClusterRow(table, "Line type", 2, _entityDashCombo, _entityDashLabel);

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
        GroupBox group = NewPropertyGroup("Label");
        TableLayoutPanel table = NewPropertyTable();
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

        AddClusterRow(table, "Colour", 0, _entityTextText, _entityTextPicker, _entityTextSwatch);

        _entityLabelPositionLabel.AutoSize = true;
        _entityLabelPositionLabel.ForeColor = SystemColors.GrayText;
        AddPropertyRow(table, "Position", _entityLabelPositionLabel, 1);

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
        GroupBox group = NewPropertyGroup("Size");
        TableLayoutPanel table = NewPropertyTable();
        group.Controls.Add(table);

        _entitySizeLabel.AutoSize = true;
        _entitySizeLabel.ForeColor = SystemColors.GrayText;
        AddPropertyRow(table, "Entity", _entitySizeLabel, 0);

        // The delineator's line width is the one built-in size the user may
        // change: it is the chart-wide DelineatorLinePt metric, which no
        // other entity reads. It is registered in the metric map under that
        // token name, so the request carries the value the user chose and
        // the dialog opens on the workbook's current width.
        _entityLineWidthNumeric.Minimum = (decimal)GanttCatalogues
            .Metrics.First(token => token.Name == "DelineatorLinePt").Minimum;
        _entityLineWidthNumeric.Maximum = (decimal)GanttCatalogues
            .Metrics.First(token => token.Name == "DelineatorLinePt").Maximum;
        _entityLineWidthNumeric.DecimalPlaces = 2;
        _entityLineWidthNumeric.Increment = 0.25m;
        _entityLineWidthNumeric.Width = 80;
        _metricNumericUpDowns["DelineatorLinePt"] = _entityLineWidthNumeric;

        _entityLineWidthLabel.AutoSize = true;
        _entityLineWidthLabel.Text = "Line width";
        AddPropertyRow(table, "Line width", _entityLineWidthNumeric, 1);

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
            _entityStyleTitleLabel.ForeColor = SystemColors.GrayText;
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
        _entityStyleTitleLabel.ForeColor = SystemColors.ControlText;
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
        _entityLabelPositionLabel.Text = DescribeLabelPosition(preset.DefaultLabelPosition);

        // The general dash, the outline width, and the bar height are
        // code-owned for a built-in style: the reader exact-matches those
        // columns, so an editable control would offer a change the workbook
        // could not keep. The delineator's line type is the exception.
        bool isDelineator = GanttStyleDashPolicy.IsUserEditable(styleKey);
        _entityDashCombo.Visible = isDelineator;
        _entityDashLabel.Visible = !isDelineator;
        _entityDashLabel.Text = DescribeStrokeDash(preset.StrokeDash);
        if (isDelineator)
        {
            PopulateDashVocabulary(_entityDashCombo);
            SelectDash(_entityDashCombo, PendingDash(styleKey, preset.StrokeDash));
        }

        _entityOutlineWidthLabel.Text = preset.StandardOutlinePt > 0
            ? $"{preset.StandardOutlinePt:0.##} pt (fixed by the built-in style)"
            : "Not applicable";

        _entitySizeLabel.Text = DescribeEntitySize(preset);

        // The delineator's line width is the one size the user may change. It
        // is the chart-wide DelineatorLinePt metric, so it is carried in the
        // metric map: the reader projects it and the writer round-trips it
        // like any other metric.
        _entityLineWidthNumeric.Enabled = isDelineator;
        _entityLineWidthNumeric.Visible = isDelineator;
        _entityLineWidthLabel.Visible = isDelineator;
    }

    /// <summary>
    /// Returns the pending line type for a style, falling back to the
    /// code-owned preset's dash when the user has not chosen one.
    /// </summary>
    /// <param name="styleKey">The built-in style key.</param>
    /// <param name="fallback">The code-owned dash to use when none is pending.</param>
    /// <returns>The pending dash, or the fallback.</returns>
    private GanttStrokeDash PendingDash(string styleKey, GanttStrokeDash fallback) =>
        _entityStyleDashes.TryGetValue(styleKey, out GanttStrokeDash pending) ? pending : fallback;

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

        // The line type is captured with the colours: the master-detail view
        // shows one style at a time, so a value the user chose must survive a
        // move to another style and back.
        if (GanttStyleDashPolicy.IsUserEditable(styleKey)
            && SelectedEntityDash() is { } dash)
        {
            _entityStyleDashes[styleKey] = dash;
        }
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
    /// Places the master column at its intended width as soon as the split is
    /// wide enough to hold it.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="SplitContainer.SplitterDistance"/> cannot be assigned in the
    /// object initializer: the split is still at its 150x100 design size there,
    /// so 240 and 250 are both clamped to 121, and
    /// <see cref="SplitContainer.FixedPanel"/> set to
    /// <see cref="FixedPanel.Panel1"/> then freezes that crushed column for the
    /// life of the dialog.
    /// </para>
    /// <para>
    /// <see cref="SplitContainer.Panel1MinSize"/> does not repair it either —
    /// the control only honours the minimum when the splitter moves, and it
    /// never moves on its own. The first layout in which the split is wide
    /// enough is therefore the moment to place the column. The handler removes
    /// itself once it has done so: dragging resizes neither the split nor its
    /// panels, so the user keeps the column wherever they leave it.
    /// </para>
    /// </remarks>
    /// <param name="split">The master/detail split.</param>
    /// <param name="masterWidth">The intended master column width.</param>
    private static void PlaceMasterColumn(SplitContainer split, int masterWidth)
    {
        void Place(object? sender, EventArgs e)
        {
            if (split.Width - split.SplitterWidth - split.Panel2MinSize < masterWidth)
            {
                return;
            }

            split.SizeChanged -= Place;
            split.SplitterDistance = masterWidth;
        }

        split.SizeChanged += Place;
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

            // Panel1MinSize is the floor the user may drag the master column
            // down to. It cannot place the column on its own — the split only
            // honours a minimum when the splitter moves, and it never moves by
            // itself — so PlaceMasterColumn below is what opens it.
            //
            // SplitterDistance must NOT be set here: the initializer runs while
            // the split is still at its 150x100 design size, so 250 clamps down
            // to 121 and FixedPanel.Panel1 then freezes that crushed width for
            // the life of the dialog.
            Panel1MinSize = 250,
            FixedPanel = FixedPanel.Panel1,
        };
        tab.Controls.Add(split);
        PlaceMasterColumn(split, 250);

        // The master: one list, grouped by family. The group heading comes
        // from the code-owned family authority, so a style can never appear
        // under a second heading and no heading is emitted twice.
        _entityStyleList.View = View.Details;
        _entityStyleList.FullRowSelect = true;
        _entityStyleList.HideSelection = false;
        _entityStyleList.HeaderStyle = ColumnHeaderStyle.None;
        _entityStyleList.Dock = DockStyle.Fill;

        // A Details-view ListView paints no rows unless it has a column, so the
        // family headings alone read as an empty list. The header is suppressed
        // (HeaderStyle.None above). The width is -2 — "fill the list's client
        // width" — so every row and its FullRowSelect highlight span the whole
        // master column at any splitter position or DPI; the earlier computed
        // pixel width (Panel1MinSize - 30) left a dead strip on the right and
        // could not follow a user drag of the splitter.
        _ = _entityStyleList.Columns.Add(string.Empty, -2);
        _entityStyleList.SelectedIndexChanged += EntityStyleList_SelectedIndexChanged;
        split.Panel1.Controls.Add(_entityStyleList);

        // The detail: the selected style's properties.
        _entityStyleDetailPanel.Dock = DockStyle.Fill;
        _entityStyleDetailPanel.AutoScroll = true;
        _entityStyleDetailPanel.Padding = new Padding(8);
        split.Panel2.Controls.Add(_entityStyleDetailPanel);

        BuildEntityStyleDetailPanel();

        foreach (GanttEntityFamily family in GanttStyleFamilies.OrderedFamilies)
        {
            GanttStylePreset[] members =
            [
                .. GanttCatalogues.StylePresets
                    .Where(preset => GanttStyleFamilies.For(preset.StyleKey) == family)
                    .Where(preset =>
                        preset.ColourCapability.HasFlag(EntityColourCapability.Fill)
                        || preset.ColourCapability.HasFlag(EntityColourCapability.Stroke)
                        || preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))

                    // The named custom activity and the named custom milestone
                    // take their style from the user, so their built-in
                    // fallback style is not something the user chose and is
                    // not offered here: those styles are created and edited on
                    // the Custom styles tab.
                    .Where(preset => !GanttUserAuthoredStyles.IsUserAuthored(preset.StyleKey)),
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
            if (TryParseStoredColour(txt.Text, out Color initial))
            {
                _colorDialog.Color = initial;
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

        if (TryParseStoredColour(text, out Color colour))
        {
            txt.BackColor = SystemColors.Window;
            UpdateSwatch(txt.Name, colour);
        }
        else
        {
            txt.BackColor = Color.FromArgb(255, 230, 230);
        }
    }

    /// <summary>
    /// Parses a colour the way the workbook stores it: exactly uppercase
    /// <c>#RRGGBB</c>.
    /// </summary>
    /// <remarks>
    /// The lenient <see cref="ColourHex"/> rule also accepts the eight-digit
    /// <c>#RRGGBBAA</c> form, which the writer and the reader reject - so the
    /// dialog accepted a value that the write then refused, and the swatch
    /// painted it with <c>ColorTranslator.FromHtml</c>, which reads eight
    /// digits as ARGB and so showed a different colour from the one typed.
    /// This is the one rule the writer, the reader and the dialog share.
    /// </remarks>
    /// <param name="text">The colour text.</param>
    /// <param name="colour">The parsed colour, or the control colour.</param>
    /// <returns>True when the text is a storable colour.</returns>
    private static bool TryParseStoredColour(string? text, out Color colour)
    {
        colour = SystemColors.Control;
        if (text is null || !GanttColourToken.IsValidHex(text))
        {
            return false;
        }

        colour = ColorTranslator.FromHtml(text);
        return true;
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
        return TryParseStoredColour(value, out Color colour)
            ? colour
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
        // MsoPatternType cannot express at all. The Delineator's own line
        // width is on the Entity styles tab, beside the delineator's other
        // properties: it belongs to that object, not to the chart.
        int y = 8;
        y = AddMetricGroup(panel, "Rows and lanes", y, "GanttRowHeightPt", "SplitterHeightPt", "SpacerHeightPt");
        y = AddMetricGroup(panel, "Grid", y, "GridLinePt", "MajorBoundaryPt");
        y = AddMetricGroup(panel, "Milestones", y, "MilestoneSizePt");
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
            Size = new Size(520, (tokens.Count * 26) + 20),
        };
        panel.Controls.Add(group);

        int innerY = 24;
        foreach ((string key, GanttMetricToken token) in tokens)
        {
            Label lbl = new()
            {
                Text = $"{DescribeMetric(key)} ({token.Minimum} to {token.Maximum} pt)",
                Location = new Point(12, innerY),
                Size = new Size(240, 20),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            group.Controls.Add(lbl);

            NumericUpDown nud = new()
            {
                Name = key,
                Location = new Point(380, innerY - 2),
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
            Size = new Size(520, (GanttChartChrome.Roles.Count * RowHeight) + 20),
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

        // The master-detail shape the Entity styles tab uses, so the two read
        // as one dialog: the list of user-created styles on the left, the
        // selected style's properties on the right.
        //
        // Every position comes from a layout container, not from a hard-coded
        // coordinate. The previous revision placed the controls absolutely
        // inside a scrolling panel, so the New button overlapped the key box,
        // the permitted-label list ran past its group, the Delete button was
        // clipped, and the content was wider than the tab, which produced a
        // horizontal scroll bar.
        SplitContainer split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,

            // See CreateColoursTab: SplitterDistance cannot be assigned here
            // (the split is still 150px wide, so 240 clamps to 121 and
            // FixedPanel keeps it there). This only floors the user's drag;
            // PlaceMasterColumn is what opens the column for the list and the
            // New row at the bottom of it.
            Panel1MinSize = 240,
            FixedPanel = FixedPanel.Panel1,
        };
        tab.Controls.Add(split);
        PlaceMasterColumn(split, 240);

        // --- The master: the list of user-created styles. ---
        Panel master = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 0, 8, 0),
        };
        split.Panel1.Controls.Add(master);

        _customStyleListGroup.Dock = DockStyle.Fill;
        _customStyleListGroup.Text = "Custom styles";
        _customStyleListGroup.Padding = new Padding(8, 4, 8, 4);

        TableLayoutPanel listLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
        };
        _ = listLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _ = listLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _ = listLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        _ = listLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _customStyleListLabel.AutoSize = true;
        listLayout.Controls.Add(_customStyleListLabel, 0, 0);

        _customStyleList.Dock = DockStyle.Fill;
        _customStyleList.IntegralHeight = false;
        _customStyleList.SelectedIndexChanged += CustomStyleList_SelectedIndexChanged;
        listLayout.Controls.Add(_customStyleList, 0, 1);

        // The empty-state hint sits under the list and wraps to the column
        // width, so the instruction is never clipped.
        _customStyleListHint.AutoSize = true;
        _customStyleListHint.MaximumSize = new Size(200, 0);
        listLayout.Controls.Add(_customStyleListHint, 0, 2);

        _customStyleListGroup.Controls.Add(listLayout);
        master.Controls.Add(_customStyleListGroup);

        // The New-style row. The key box and the button sit in a two-column
        // table: the key box is the flexible column and the button owns an
        // autosized one, so the button can never be pushed past the master
        // column's edge. A FlowLayoutPanel could not make that guarantee — it
        // laid the button out at x=138 and clipped it at the column's 113px.
        TableLayoutPanel newRow = new TableLayoutPanel
        {
            Dock = DockStyle.Bottom,
            ColumnCount = 2,
            RowCount = 1,
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0),
        };
        _ = newRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _ = newRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        _ = newRow.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // Anchored rather than docked: a Docked child measures as empty to an
        // AutoSize row (the same collapse the property groups hit), whereas the
        // anchor stretches the box across its cell while still reporting its
        // own 22px height for the row to grow around.
        _customNewKeyText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        _customNewKeyText.Margin = new Padding(0, 0, 8, 0);
        newRow.Controls.Add(_customNewKeyText, 0, 0);
        _customNewButton.Click += CustomNewButton_Click;
        _customNewButton.Margin = new Padding(0);
        newRow.Controls.Add(_customNewButton, 1, 0);
        master.Controls.Add(newRow);

        // --- The detail: the selected style's properties. ---
        _customStyleEditGroup.Dock = DockStyle.Fill;
        _customStyleEditGroup.Text = "Edit selected style";
        _customStyleEditGroup.Padding = new Padding(8);
        split.Panel2.Controls.Add(_customStyleEditGroup);

        // The property groups are stacked in a table rather than a flow panel:
        // a flow panel gives every child its own preferred width, so the groups
        // were 420px wide inside a 616px panel and left a wide empty strip on
        // the right. A table's single percent-width column gives every group
        // the panel's full width, and each group still hugs its own content.
        //
        // The table is two rows: the properties scroll, and the Delete row is
        // pinned below them so it can never be pushed out of sight by the
        // groups above it.
        TableLayoutPanel detail = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
        };
        _ = detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _ = detail.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        _ = detail.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _customStyleEditGroup.Controls.Add(detail);

        TableLayoutPanel properties = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            AutoScroll = true,
        };
        _ = properties.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        for (int row = 0; row < 5; row++)
        {
            _ = properties.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        }

        // The style key, read-only: it names the row, and a rename is a
        // delete plus a create, so the key is not edited in place.
        TableLayoutPanel keyRow = NewPropertyTable();
        keyRow.Dock = DockStyle.Top;
        _customStyleKeyText.ReadOnly = true;
        AddPropertyRow(keyRow, _customStyleKeyLabel, _customStyleKeyText, 0);
        properties.Controls.Add(keyRow, 0, 0);

        // The property groups, one per concern: Fill, Outline, Label, Size.
        TableLayoutPanel fillTable = NewPropertyTable();
        AddCustomColourRow(fillTable, "Colour", _customStyleFillText, _customFillColorButton, _customFillSwatch, 0);
        AddCustomPropertyRow(fillTable, "Pattern", _customHatchPatternCombo, 1);
        GroupBox fillGroup = NewPropertyGroup("Fill");
        fillGroup.Controls.Add(fillTable);
        properties.Controls.Add(fillGroup, 0, 1);

        TableLayoutPanel outlineTable = NewPropertyTable();
        AddCustomColourRow(outlineTable, "Colour", _customStyleStrokeText, _customStrokeColorButton, _customStrokeSwatch, 0);
        AddCustomPropertyRow(outlineTable, "Width", _customOutlineWidthNumeric, 1);
        AddCustomPropertyRow(outlineTable, "Line style", _customStrokeDashCombo, 2);
        GroupBox outlineGroup = NewPropertyGroup("Outline");
        outlineGroup.Controls.Add(outlineTable);
        properties.Controls.Add(outlineGroup, 0, 2);

        TableLayoutPanel labelTable = NewPropertyTable();
        AddCustomColourRow(labelTable, "Colour", _customStyleTextText, _customTextColorButton, _customTextSwatch, 0);
        AddCustomPropertyRow(labelTable, "Default", _customDefaultLabelCombo, 1);
        AddCustomPropertyRow(labelTable, "Allowed", _customAllowedLabelList, 2);
        GroupBox labelGroup = NewPropertyGroup("Label");
        labelGroup.Controls.Add(labelTable);
        properties.Controls.Add(labelGroup, 0, 3);

        TableLayoutPanel sizeTable = NewPropertyTable();
        AddCustomPropertyRow(sizeTable, "Bar height", _customActivityHeightNumeric, 0);
        AddCustomPropertyRow(sizeTable, "Milestone size", _customMilestoneSizeNumeric, 1);
        _customMilestoneSizeNumeric.Minimum = 0;
        _customMilestoneSizeNumeric.AccessibleDescription =
            "Zero uses the chart's milestone size.";
        _customActivityHeightNumeric.AccessibleDescription =
            "Zero uses the height this entity kind's own style carries.";
        GroupBox sizeGroup = NewPropertyGroup("Size");
        sizeGroup.Controls.Add(sizeTable);
        properties.Controls.Add(sizeGroup, 0, 4);

        detail.Controls.Add(properties, 0, 0);

        // Delete only applies to a selected style, so it gets its own row at
        // the foot of the property panel rather than sitting over a group.
        FlowLayoutPanel deleteRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = false,
            Margin = new Padding(0),
        };
        _customDeleteButton.Click += CustomDeleteButton_Click;
        _customDeleteButton.Margin = new Padding(0);
        deleteRow.Controls.Add(_customDeleteButton);
        detail.Controls.Add(deleteRow, 0, 1);

        PopulateCustomStyleVocabularies();

        // The permitted-label list's window is font-relative rather than a
        // fixed pixel height, so it always shows the same whole number of
        // positions and never a sliver of a twelfth. It is sized to hold every
        // position in the vocabulary, so the list never has to be scrolled to
        // reach a position. With IntegralHeight on, the control snaps the
        // requested height down to whole items.
        _customAllowedLabelList.Height =
            (_customAllowedLabelList.ItemHeight * _customAllowedLabelList.Items.Count) + 6;

        UpdateCustomStyleListHint();
#pragma warning restore CA1303
    }

    /// <summary>
    /// Adds a labelled colour row (hex text box, picker button, swatch) to a
    /// property group's table.
    /// </summary>
    /// <param name="table">The group's property table.</param>
    /// <param name="label">The property label.</param>
    /// <param name="textBox">The hex text box.</param>
    /// <param name="pickerButton">The colour picker button.</param>
    /// <param name="swatch">The colour swatch.</param>
    /// <param name="row">The row the control sits in.</param>
    private void AddCustomColourRow(
        TableLayoutPanel table,
        string label,
        TextBox textBox,
        Button pickerButton,
        Panel swatch,
        int row)
    {
        textBox.TextChanged += CustomStyleColourTextChanged;
        pickerButton.Click += CustomFillColorButton_Click;

        // The three colour controls travel together in one right-anchored
        // cluster, so the swatch can never drift away from the colour it
        // shows — which is what made the swatch look like it belonged to a
        // different property.
        AddClusterRow(table, label, row, textBox, pickerButton, swatch);
    }

    /// <summary>
    /// Adds a labelled row holding a single property control.
    /// </summary>
    /// <param name="table">The group's property table.</param>
    /// <param name="label">The property label.</param>
    /// <param name="control">The property control.</param>
    /// <param name="row">The row the control sits in.</param>
    private static void AddCustomPropertyRow(
        TableLayoutPanel table,
        string label,
        Control control,
        int row) => AddPropertyRow(table, label, control, row);

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
            _ = _customAllowedLabelList.Items.Add(new LabelPositionItem(position, DescribeLabelPosition(position)));
        }

        // The default-position combo is NOT populated here: it offers exactly
        // the positions the style PERMITS, so it is rebuilt from the checked
        // items of the permitted-label list (R5.6d WI-C). The writer's rule
        // is that the default must be one of the permitted positions, so a
        // combo that offered every position would let the user choose a pair
        // the workbook cannot store.
        _ = _customDefaultLabelCombo.Items.Add(
            new LabelPositionItem(GanttLabelPosition.Auto, DescribeLabelPosition(GanttLabelPosition.Auto)));

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
    internal sealed record LabelPositionItem(GanttLabelPosition Position, string Label)
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

        for (int index = 0; index < _customAllowedLabelList.Items.Count; index++)
        {
            bool isChecked = _customAllowedLabelList.Items[index] is LabelPositionItem item
                && state.AllowedLabelPositions.Contains(item.Position);
            _customAllowedLabelList.SetItemChecked(index, isChecked);
        }

        // The default combo offers exactly the permitted positions, so it is
        // rebuilt from the list before the stored default is selected (R5.6d
        // WI-C). A stored default that the style does not permit falls back
        // to a permitted one, because the workbook cannot store the pair.
        RebuildDefaultLabelCombo(CheckedAllowedLabelPositions());
        _customDefaultLabelCombo.SelectedItem =
            _customDefaultLabelCombo.Items.Cast<LabelPositionItem>()
                .FirstOrDefault(item => item.Position == state.DefaultLabelPosition);

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
        // label positions a span or a milestone both accept. The milestone
        // size starts at ZERO, which means "use the chart's milestone size":
        // it is an override, not a second source of truth, so the chart-wide
        // knob stays live for a style that does not set it.
        //
        // The BAR height starts at zero for the same reason, and for the same
        // contract: zero is "not applicable to this entity kind", so a style
        // that sets no height of its own inherits the height its row's own
        // default style carries (R5.6a). Seeding the code-owned default
        // instead would pin a copied value into the style, and the row would
        // then keep that number even if the kind's own height changed.
        _customStyleEdits[newKey] = new CustomStyleState(
            "#FFFFFF",
            "#000000",
            "#000000",
            GanttHatchPattern.None,
            GanttCatalogues.MetricDefault("StandardOutlinePt"),
            0,
            0,
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
        _colorDialog.Color = TryParseStoredColour(textBox.Text, out Color initial)
            ? initial
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
    /// Rebuilds the default-position combo so it offers exactly the positions
    /// the style permits.
    /// </summary>
    /// <remarks>
    /// R5.6d WI-C: the writer's rule is that the default label position must
    /// be one of the permitted positions, and the reader enforces it. A combo
    /// that offered every position let the user choose a pair the workbook
    /// cannot store, which the write then refused. A default that is no
    /// longer permitted moves to the first position that is, so the edit
    /// never carries an inconsistent pair.
    /// </remarks>
    /// <param name="permitted">The permitted positions.</param>
    private void RebuildDefaultLabelCombo(HashSet<GanttLabelPosition> permitted)
    {
        GanttLabelPosition? current = _customDefaultLabelCombo.SelectedItem is LabelPositionItem selected
            ? selected.Position
            : null;

        _customDefaultLabelCombo.Items.Clear();
        foreach (GanttLabelPosition position in Enum.GetValues<GanttLabelPosition>())
        {
            if (permitted.Contains(position))
            {
                _ = _customDefaultLabelCombo.Items.Add(
                    new LabelPositionItem(position, DescribeLabelPosition(position)));
            }
        }

        List<LabelPositionItem> items = [.. _customDefaultLabelCombo.Items.Cast<LabelPositionItem>()];
        LabelPositionItem? chosen =
            items.FirstOrDefault(item => current is { } position && item.Position == position)
            ?? items.FirstOrDefault();
        _customDefaultLabelCombo.SelectedItem = chosen;
    }

    /// <summary>
    /// Returns the permitted label positions the user has ticked in the
    /// permitted-label list.
    /// </summary>
    /// <returns>The permitted positions.</returns>
    private HashSet<GanttLabelPosition> CheckedAllowedLabelPositions()
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

        // The default follows the permitted set (R5.6d WI-C), so the combo is
        // rebuilt FIRST and the state is then written from the corrected
        // combo. The rebuild's own change event is suppressed, because the
        // list's check state has not changed yet at this point.
        bool wasSuppressed = _suppressCustomStyleSync;
        _suppressCustomStyleSync = true;
        try
        {
            RebuildDefaultLabelCombo(selected);
        }
        finally
        {
            _suppressCustomStyleSync = wasSuppressed;
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

        // The line types. Only the delineator's is user-editable, so only that
        // one carries a value; every other built-in style's dash is
        // code-owned and the reader exact-matches it against the preset.
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            _entityStyleDashes[preset.StyleKey] = preset.StrokeDash;
        }

        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            if (edit.StrokeDash is { } dash && GanttStyleDashPolicy.IsUserEditable(edit.StyleKey))
            {
                _entityStyleDashes[edit.StyleKey] = dash;
            }
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
        // Validate all colour text boxes. Surrounding whitespace is trimmed
        // first: the strict rule is exactly seven characters, so a stray
        // space would otherwise be refused where the user meant the colour
        // they typed - and the trimmed value is what the workbook stores.
        foreach (KeyValuePair<string, TextBox> kvp in _colorTextBoxes)
        {
            kvp.Value.Text = kvp.Value.Text.Trim();
            if (!string.IsNullOrEmpty(kvp.Value.Text) && !TryParseStoredColour(kvp.Value.Text, out _))
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
        _customStyleFillText.Text = _customStyleFillText.Text.Trim();
        _customStyleStrokeText.Text = _customStyleStrokeText.Text.Trim();
        _customStyleTextText.Text = _customStyleTextText.Text.Trim();
        if (!string.IsNullOrEmpty(_customStyleFillText.Text) && !TryParseStoredColour(_customStyleFillText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style fill colour '" + _customStyleFillText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleFillText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleStrokeText.Text) && !TryParseStoredColour(_customStyleStrokeText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style stroke colour '" + _customStyleStrokeText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[2];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleStrokeText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleTextText.Text) && !TryParseStoredColour(_customStyleTextText.Text, out _))
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
            kvp.Value.Text = kvp.Value.Text.Trim();
            if (!string.IsNullOrEmpty(kvp.Value.Text) && !TryParseStoredColour(kvp.Value.Text, out _))
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

        // A custom style must permit at least one label position, and its
        // default must be one of them: the writer refuses that pair, so it is
        // refused here with the list focused rather than after the write
        // (R5.6d WI-C).
        foreach (KeyValuePair<string, CustomStyleState> edit in _customStyleEdits)
        {
            if (edit.Value.AllowedLabelPositions.Count == 0)
            {
                _warningLabel.Text = "Custom style '" + edit.Key
                    + "' must permit at least one label position.";
                _warningLabel.Visible = true;
                _tabControl.SelectedTab = _tabControl.TabPages[2];
                _customStyleEditGroup.Enabled = true;
                _ = _customAllowedLabelList.Focus();
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
                GanttStrokeDash? dash = GanttStyleDashPolicy.IsUserEditable(preset.StyleKey)
                    ? PendingDash(preset.StyleKey, preset.StrokeDash)
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

        ApplyCodeOwnedDefaults();
    }

    /// <summary>
    /// Restores every code-owned default the dialog can edit: the built-in
    /// styles' colours and line types, the chart-chrome colours, and the
    /// chart-wide metrics. Runs on the STA thread; never throws.
    /// </summary>
    /// <remarks>
    /// The reset must put the PRESET back, not blank the colour: an empty
    /// colour means "no fill" to the writer, so the earlier behaviour removed
    /// the colour instead of restoring it.
    /// </remarks>
    private void ApplyCodeOwnedDefaults()
    {
        // Restore the code-owned preset values for every built-in style.
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            _entityStyleColours[preset.StyleKey] = new EntityStyleColours(
                preset.FillColour,
                preset.StrokeColour,
                preset.TextColour);
            _entityStyleDashes[preset.StyleKey] = preset.StrokeDash;
        }

        // The chart-chrome colours go back to the catalogue defaults too, so
        // a reset really does return the chart to the shipped look.
        foreach (GanttChartChromeRole role in GanttChartChrome.Roles)
        {
            string key = GanttChartChrome.SettingKeyFor(role);
            if (_chromeColourTextBoxes.TryGetValue(key, out TextBox? chromeText))
            {
                chromeText.Text = GanttChartChrome.DefaultHexFor(role);
                UpdateSwatch(key, ParseColourOrDefault(chromeText.Text));
            }
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

        // The detail panel shows one style at a time, so it is refreshed from
        // the restored values.
        if (_selectedEntityStyleKey is { } selectedKey)
        {
            ShowEntityStyle(selectedKey);
        }

        _warningLabel.Visible = false;
    }
}
