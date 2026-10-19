using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// WinForms modal dialog for R5.6 theme settings (Colours / Heights and sizes / Grid and bands tabs).
/// Runs on the Excel main STA thread via ShowDialog; fully offline.
/// </summary>
/// <summary>Snapshot of one custom style's colour state.</summary>
internal sealed class CustomStyleState(string fill, string stroke, string text)
{
    public string Fill { get; } = fill;
    public string Stroke { get; } = stroke;
    public string Text { get; } = text;
}

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
    private readonly Dictionary<string, Button> _colorPickerButtons = [];
    private readonly Dictionary<string, NumericUpDown> _metricNumericUpDowns = [];
    private readonly Dictionary<string, Label> _metricLabels = [];

    // Custom styles tab controls
    private const string _customStyleKeyLabel = "Style key:";
    private const string _customStyleFillLabel = "Fill";
    private const string _customStyleStrokeLabel = "Stroke";
    private const string _customStyleTextLabel = "Text";
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
        ClientSize = new Size(560, 480);
        AutoScaleMode = AutoScaleMode.Dpi;

        Controls.Add(_tabControl);

        CreateColoursTab();
        CreateHeightsTab();
        CreateGridTab();
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

    private void CreateColoursTab()
    {
        TabPage tab = new TabPage("Colours")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        tab.Controls.Add(panel);

        int y = 8;
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            if (!preset.ColourCapability.HasFlag(EntityColourCapability.Fill) &&
                !preset.ColourCapability.HasFlag(EntityColourCapability.Stroke) &&
                !preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                continue;
            }

            GroupBox group = new GroupBox
            {
                Text = preset.DisplayName,
                Location = new Point(8, y),
                Size = new Size(500, 100),
                AutoSize = true,
            };
            panel.Controls.Add(group);

            int innerY = 24;

            if (preset.ColourCapability.HasFlag(EntityColourCapability.Fill) ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                AddColourRow(group, $"fill_{preset.StyleKey}", "Fill", preset.FillColour, innerY);
                innerY += 28;
            }

            if (preset.ColourCapability.HasFlag(EntityColourCapability.Stroke) ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                AddColourRow(group, $"stroke_{preset.StyleKey}", "Stroke", preset.StrokeColour, innerY);
                innerY += 28;
            }

            if (preset.TextColour.Length > 0 ||
                preset.ColourCapability.HasFlag(EntityColourCapability.StyleDefined))
            {
                AddColourRow(group, $"text_{preset.StyleKey}", "Text", preset.TextColour, innerY);
                innerY += 28;
            }

            y += group.Height + 8;
        }
    }

    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void AddColourRow(Control parent, string key, string label, string initialValue, int top)
    {
        Label lbl = new Label
        {
            Text = label,
            Location = new Point(12, top + 3),
            Size = new Size(50, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        parent.Controls.Add(lbl);

        TextBox txt = new TextBox
        {
            Name = key,
            Location = new Point(70, top),
            Size = new Size(100, 22),
            Text = initialValue,
            CharacterCasing = CharacterCasing.Upper,
            MaxLength = 9,
        };
        txt.Leave += (_, _) => ValidateColourTextBox(txt);
        parent.Controls.Add(txt);
        _colorTextBoxes[key] = txt;

        Button btn = new Button
        {
            Text = "\u2026",
            Location = new Point(176, top - 1),
            Size = new Size(26, 24),
            Tag = key,
        };
        btn.Click += ColorPickerButton_Click;
        parent.Controls.Add(btn);
        _colorPickerButtons[key] = btn;

        Panel swatch = new Panel
        {
            Location = new Point(210, top),
            Size = new Size(24, 22),
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = ParseColourOrDefault(initialValue),
            Tag = key,
        };
        parent.Controls.Add(swatch);
    }

    private void ColorPickerButton_Click(object? sender, EventArgs e)
    {
        if (sender is Button btn && btn.Tag is string key &&
            _colorTextBoxes.TryGetValue(key, out TextBox? txt))
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
        foreach (Control ctrl in Controls)
        {
            if (ctrl is Panel p && p.Tag?.ToString() == key)
            {
                p.BackColor = color;
                break;
            }
        }
    }

    private static Color ParseColourOrDefault(string value)
    {
        return ColourHex.TryParse(value, out ColourHex? c) && c is not null
            ? Color.FromArgb((int)c.ARGB)
            : SystemColors.Control;
    }

    private void CreateHeightsTab()
    {
        TabPage tab = new TabPage("Heights & sizes")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        tab.Controls.Add(panel);

        int y = 8;
        string[] metricKeys =
        [
            "ActivityHeightPt",
            "MilestoneSizePt",
            "StandardOutlinePt",
            "GanttRowHeightPt",
            "SplitterHeightPt",
            "SpacerHeightPt",
        ];

        foreach (string key in metricKeys)
        {
            GanttMetricToken? token = GanttCatalogues.Metrics.FirstOrDefault(t => t.Name == key);
            if (token == null)
            {
                continue;
            }

            GroupBox group = new GroupBox
            {
                Text = key,
                Location = new Point(8, y),
                Size = new Size(500, 60),
                AutoSize = true,
            };
            panel.Controls.Add(group);

            Label lbl = new Label
            {
                Text = $"Range: {token.Minimum} – {token.Maximum} pt",
                Location = new Point(12, 24),
                AutoSize = true,
            };
            group.Controls.Add(lbl);

            NumericUpDown nud = new NumericUpDown
            {
                Name = key,
                Location = new Point(12, 46),
                Size = new Size(80, 22),
                Minimum = (decimal)token.Minimum,
                Maximum = (decimal)token.Maximum,
                DecimalPlaces = 2,
                Increment = 0.25m,
            };
            group.Controls.Add(nud);
            _metricNumericUpDowns[key] = nud;
            _metricLabels[key] = lbl;

            y += group.Height + 8;
        }
    }

    private void CreateGridTab()
    {
        TabPage tab = new TabPage("Grid & bands")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        Panel panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
        tab.Controls.Add(panel);

        int y = 8;
        string[] metricKeys =
        [
            "GridLinePt",
            "MajorBoundaryPt",
            "DelineatorLinePt",
            "HatchPitchPt",
            "HatchLinePt",
        ];

        foreach (string key in metricKeys)
        {
            GanttMetricToken? token = GanttCatalogues.Metrics.FirstOrDefault(t => t.Name == key);
            if (token == null)
            {
                continue;
            }

            GroupBox group = new GroupBox
            {
                Text = key,
                Location = new Point(8, y),
                Size = new Size(500, 60),
                AutoSize = true,
            };
            panel.Controls.Add(group);

            Label lbl = new Label
            {
                Text = $"Range: {token.Minimum} – {token.Maximum} pt",
                Location = new Point(12, 24),
                AutoSize = true,
            };
            group.Controls.Add(lbl);

            NumericUpDown nud = new NumericUpDown
            {
                Name = key,
                Location = new Point(12, 46),
                Size = new Size(80, 22),
                Minimum = (decimal)token.Minimum,
                Maximum = (decimal)token.Maximum,
                DecimalPlaces = 2,
                Increment = 0.25m,
            };
            group.Controls.Add(nud);
            _metricNumericUpDowns[key] = nud;
            _metricLabels[key] = lbl;

            y += group.Height + 8;
        }
    }

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

        // Left: list of existing custom styles
        _customStyleListGroup.Location = new Point(8, 8);
        _customStyleListLabel.Location = new Point(16, 24);
        _customStyleList.Location = new Point(16, 48);

        _customStyleList.SelectedIndexChanged += CustomStyleList_SelectedIndexChanged;
        _customStyleListGroup.Controls.Add(_customStyleListLabel);
        _customStyleListGroup.Controls.Add(_customStyleList);
        panel.Controls.Add(_customStyleListGroup);

        // Empty-state hint (full width so the instruction stays readable).
        _customStyleListHint.AutoSize = true;
        _customStyleListHint.Location = new Point(8, 272);
        panel.Controls.Add(_customStyleListHint);

        // Right: edit panel for the selected custom style
        _customStyleEditGroup.Location = new Point(240, 8);

        Label keyLbl = new Label
        {
            Text = _customStyleKeyLabel,
            Location = new Point(16, 28),
            Size = new Size(64, 20),
        };
        _customStyleKeyText.Location = new Point(84, 24);

        Label fillLbl = new Label
        {
            Text = _customStyleFillLabel,
            Location = new Point(16, 64),
            Size = new Size(32, 20),
        };
        _customStyleFillText.Location = new Point(84, 60);
        _customFillColorButton.Location = new Point(192, 58);
        _customFillSwatch.Location = new Point(224, 60);
        _customFillColorButton.Click += CustomFillColorButton_Click;
        _customStyleFillText.TextChanged += CustomStyleColourTextChanged;

        Label strokeLbl = new Label
        {
            Text = _customStyleStrokeLabel,
            Location = new Point(16, 96),
            Size = new Size(44, 20),
        };
        _customStyleStrokeText.Location = new Point(84, 92);
        _customStrokeColorButton.Location = new Point(192, 90);
        _customStrokeSwatch.Location = new Point(224, 92);
        _customStrokeColorButton.Click += CustomStrokeColorButton_Click;
        _customStyleStrokeText.TextChanged += CustomStyleColourTextChanged;

        Label textLbl = new Label
        {
            Text = _customStyleTextLabel,
            Location = new Point(16, 128),
            Size = new Size(36, 20),
        };
        _customStyleTextText.Location = new Point(84, 124);
        _customTextColorButton.Location = new Point(192, 122);
        _customTextSwatch.Location = new Point(224, 124);
        _customTextColorButton.Click += CustomTextColorButton_Click;
        _customStyleTextText.TextChanged += CustomStyleColourTextChanged;

        // New-style entry is always available (not gated behind a selection).
        _customNewButton.Click += CustomNewButton_Click;

        // Delete only applies to a selected style; it lives inside the edit group.
        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Location = new Point(16, 172),
            Size = new Size(280, 30),
            FlowDirection = FlowDirection.LeftToRight,
        };
        _customDeleteButton.Click += CustomDeleteButton_Click;
        buttonPanel.Controls.Add(_customDeleteButton);

        _customStyleEditGroup.Controls.Add(keyLbl);
        _customStyleEditGroup.Controls.Add(_customStyleKeyText);
        _customStyleEditGroup.Controls.Add(fillLbl);
        _customStyleEditGroup.Controls.Add(_customStyleFillText);
        _customStyleEditGroup.Controls.Add(_customFillColorButton);
        _customStyleEditGroup.Controls.Add(_customFillSwatch);
        _customStyleEditGroup.Controls.Add(strokeLbl);
        _customStyleEditGroup.Controls.Add(_customStyleStrokeText);
        _customStyleEditGroup.Controls.Add(_customStrokeColorButton);
        _customStyleEditGroup.Controls.Add(_customStrokeSwatch);
        _customStyleEditGroup.Controls.Add(textLbl);
        _customStyleEditGroup.Controls.Add(_customStyleTextText);
        _customStyleEditGroup.Controls.Add(_customTextColorButton);
        _customStyleEditGroup.Controls.Add(_customTextSwatch);
        _customStyleEditGroup.Controls.Add(buttonPanel);
        panel.Controls.Add(_customStyleEditGroup);

        // Always-visible new-style row: key entry + New button.
        FlowLayoutPanel customNewRow = new FlowLayoutPanel
        {
            Location = new Point(8, 296),
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
        };
        Label newKeyLabel = new Label
        {
            Text = "New key:",
            AutoSize = true,
        };
        customNewRow.Controls.Add(newKeyLabel);
        customNewRow.Controls.Add(_customNewKeyText);
        customNewRow.Controls.Add(_customNewButton);
        panel.Controls.Add(customNewRow);

        UpdateCustomStyleListHint();
#pragma warning restore CA1303
    }

    private void CustomStyleList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        _selectedCustomStyleKey = _customStyleList.SelectedItem as string;
        bool hasSelection = _selectedCustomStyleKey is not null;
        _customDeleteButton.Enabled = hasSelection;
        _customStyleEditGroup.Enabled = hasSelection;

        if (hasSelection && _selectedCustomStyleKey is { } key &&
            _customStyleEdits.TryGetValue(key, out CustomStyleState? colours))
        {
            _customStyleKeyText.Text = _selectedCustomStyleKey;
            _customStyleFillText.Text = colours.Fill;
            _customStyleStrokeText.Text = colours.Stroke;
            _customStyleTextText.Text = colours.Text;
            _customFillSwatch.BackColor = ParseColourOrDefault(colours.Fill);
            _customStrokeSwatch.BackColor = ParseColourOrDefault(colours.Stroke);
            _customTextSwatch.BackColor = ParseColourOrDefault(colours.Text);
        }
        else
        {
            ClearCustomStyleEditor();
        }

        UpdateCustomStyleListHint();
    }

    private void ClearCustomStyleEditor()
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
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            _ = _customNewKeyText.Focus();
            return;
        }

        if (IsBuiltInCustomKey(newKey))
        {
            _warningLabel.Text = $"Style key '{newKey}' is a built-in preset and cannot be used as a custom style.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            return;
        }

        if (_customStyleEdits.ContainsKey(newKey))
        {
            _warningLabel.Text = $"Custom style '{newKey}' already exists.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            return;
        }

        _warningLabel.Visible = false;

        _customStyleEdits[newKey] = new CustomStyleState("#FFFFFF", "#000000", "#000000");
        _ = _customStyleList.Items.Add(newKey);
        _customStyleList.SelectedItem = newKey;
    }
#pragma warning restore CA1303

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
        if (_selectedCustomStyleKey is null)
        {
            return;
        }

        _customStyleEdits[_selectedCustomStyleKey] = new CustomStyleState(
            _customStyleFillText.Text,
            _customStyleStrokeText.Text,
            _customStyleTextText.Text);
    }

    private void PopulateFromRequest(ThemeEditRequest request)
    {
        // Colours
        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            if (_colorTextBoxes.TryGetValue($"fill_{edit.StyleKey}", out TextBox? fillTxt))
            {
                fillTxt.Text = edit.FillColour;
            }
            if (_colorTextBoxes.TryGetValue($"stroke_{edit.StyleKey}", out TextBox? strokeTxt))
            {
                strokeTxt.Text = edit.StrokeColour;
            }
            if (_colorTextBoxes.TryGetValue($"text_{edit.StyleKey}", out TextBox? textTxt))
            {
                textTxt.Text = edit.TextColour;
            }

            // Update swatches
            UpdateSwatch($"fill_{edit.StyleKey}", ParseColourOrDefault(edit.FillColour));
            UpdateSwatch($"stroke_{edit.StyleKey}", ParseColourOrDefault(edit.StrokeColour));
            UpdateSwatch($"text_{edit.StyleKey}", ParseColourOrDefault(edit.TextColour));
        }

        // Metrics
        foreach (KeyValuePair<string, double> kvp in request.MetricEdits)
        {
            if (_metricNumericUpDowns.TryGetValue(kvp.Key, out NumericUpDown? nud))
            {
                nud.Value = (decimal)kvp.Value;
            }
        }

        // Custom styles: load saved styles into the edit dictionary and list so
        // they are visible on the tab and are not spuriously deleted on OK.
        foreach (CustomStyleDef def in request.CustomStyleDefs ?? [])
        {
            _customStyleEdits[def.StyleKey] = new CustomStyleState(def.FillColour, def.StrokeColour, def.TextColour);
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
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleFillText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleStrokeText.Text) && !ColourHex.TryParse(_customStyleStrokeText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style stroke colour '" + _customStyleStrokeText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleStrokeText.Focus();
            DialogResult = DialogResult.None;
            return;
        }
        if (!string.IsNullOrEmpty(_customStyleTextText.Text) && !ColourHex.TryParse(_customStyleTextText.Text, out _))
        {
            _warningLabel.Text = "Invalid custom style text colour '" + _customStyleTextText.Text + "'. Use uppercase #RRGGBB.";
            _warningLabel.Visible = true;
            _tabControl.SelectedTab = _tabControl.TabPages[3];
            _customStyleEditGroup.Enabled = true;
            _ = _customStyleTextText.Focus();
            DialogResult = DialogResult.None;
            return;
        }

        Dictionary<string, double> metricEdits = [];
        foreach (KeyValuePair<string, NumericUpDown> kvp in _metricNumericUpDowns)
        {
            metricEdits[kvp.Key] = (double)kvp.Value.Value;
        }

        // Build request
        List<ThemeStyleEdit> styleEdits = [];
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            string fill = _colorTextBoxes.TryGetValue($"fill_{preset.StyleKey}", out TextBox? ft) ? ft.Text : "";
            string stroke = _colorTextBoxes.TryGetValue($"stroke_{preset.StyleKey}", out TextBox? st) ? st.Text : "";
            string text = _colorTextBoxes.TryGetValue($"text_{preset.StyleKey}", out TextBox? tt) ? tt.Text : "";

            if (fill.Length > 0 || stroke.Length > 0 || text.Length > 0)
            {
                styleEdits.Add(new ThemeStyleEdit(preset.StyleKey, fill, stroke, text));
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
                    TextColour: state.Text));
            }
            else if (original.FillColour != state.Fill
                || original.StrokeColour != state.Stroke
                || original.TextColour != state.Text)
            {
                // Colour-only change on an existing key is an upsert by key:
                // the key is unchanged, so it can never collide with another
                // key (R5.6a D1). A Rename with identical old/new key would
                // falsely fail the request's duplicate-key check.
                customStyleEdits.Add(new CustomStyleEdit(
                    CustomStyleOperation.Update,
                    key,
                    DisplayName: key,
                    FillColour: state.Fill,
                    StrokeColour: state.Stroke,
                    TextColour: state.Text));
            }
        }

        ThemeEditRequest request = new ThemeEditRequest(styleEdits, metricEdits, customStyleEdits);

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
