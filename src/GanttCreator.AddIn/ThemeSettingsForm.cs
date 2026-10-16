using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// WinForms modal dialog for R5.6 theme settings (Colours / Heights and sizes / Grid and bands tabs).
/// Runs on the Excel main STA thread via ShowDialog; fully offline.
/// </summary>
internal sealed partial class ThemeSettingsForm : Form
{
    private ThemeEditRequest? _validatedRequest;

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
        }
        base.Dispose(disposing);
    }

    /// <summary>
    /// Gets the validated request if the user clicked OK, or null if cancelled/invalid.
    /// </summary>
    internal ThemeEditRequest? GetValidatedRequest() => _validatedRequest;

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

        // Bottom buttons + warning label
        var bottomPanel = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 60,
            Padding = new Padding(8),
        };
        Controls.Add(bottomPanel);

        _warningLabel.Font = new Font(Font.FontFamily, 8.25f, FontStyle.Regular);
        bottomPanel.Controls.Add(_warningLabel);

        var buttonPanel = new FlowLayoutPanel
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
        var tab = new TabPage("Colours")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
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

            var group = new GroupBox
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
        var lbl = new Label
        {
            Text = label,
            Location = new Point(12, top + 3),
            Size = new Size(50, 20),
            TextAlign = ContentAlignment.MiddleLeft,
        };
        parent.Controls.Add(lbl);

        var txt = new TextBox
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

        var btn = new Button
        {
            Text = "\u2026",
            Location = new Point(176, top - 1),
            Size = new Size(26, 24),
            Tag = key,
        };
        btn.Click += ColorPickerButton_Click;
        parent.Controls.Add(btn);
        _colorPickerButtons[key] = btn;

        var swatch = new Panel
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
            _colorTextBoxes.TryGetValue(key, out var txt))
        {
            if (ColourHex.TryParse(txt.Text, out ColourHex? initial))
            {
                _colorDialog.Color = System.Drawing.Color.FromArgb((int)initial!.ARGB);
            }

            if (_colorDialog.ShowDialog(this) == DialogResult.OK)
            {
                var color = _colorDialog.Color;
                var hex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
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
            UpdateSwatch(txt.Name, System.Drawing.ColorTranslator.FromHtml(text));
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
        return ColourHex.TryParse(value, out var c) && c is not null
            ? System.Drawing.Color.FromArgb((int)c.ARGB)
            : SystemColors.Control;
    }

    private void CreateHeightsTab()
    {
        var tab = new TabPage("Heights & sizes")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
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
            var token = GanttCatalogues.Metrics.FirstOrDefault(t => t.Name == key);
            if (token == null)
            {
                continue;
            }

            var group = new GroupBox
            {
                Text = key,
                Location = new Point(8, y),
                Size = new Size(500, 60),
                AutoSize = true,
            };
            panel.Controls.Add(group);

            var lbl = new Label
            {
                Text = $"Range: {token.Minimum} – {token.Maximum} pt",
                Location = new Point(12, 24),
                AutoSize = true,
            };
            group.Controls.Add(lbl);

            var nud = new NumericUpDown
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
        var tab = new TabPage("Grid & bands")
        {
            Padding = new Padding(8),
        };
        _tabControl.TabPages.Add(tab);

        var panel = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
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
            var token = GanttCatalogues.Metrics.FirstOrDefault(t => t.Name == key);
            if (token == null)
            {
                continue;
            }

            var group = new GroupBox
            {
                Text = key,
                Location = new Point(8, y),
                Size = new Size(500, 60),
                AutoSize = true,
            };
            panel.Controls.Add(group);

            var lbl = new Label
            {
                Text = $"Range: {token.Minimum} – {token.Maximum} pt",
                Location = new Point(12, 24),
                AutoSize = true,
            };
            group.Controls.Add(lbl);

            var nud = new NumericUpDown
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

    private void PopulateFromRequest(ThemeEditRequest request)
    {
        // Colours
        foreach (var edit in request.StyleEdits)
        {
            if (_colorTextBoxes.TryGetValue($"fill_{edit.StyleKey}", out var fillTxt))
            {
                fillTxt.Text = edit.FillColour;
            }
            if (_colorTextBoxes.TryGetValue($"stroke_{edit.StyleKey}", out var strokeTxt))
            {
                strokeTxt.Text = edit.StrokeColour;
            }
            if (_colorTextBoxes.TryGetValue($"text_{edit.StyleKey}", out var textTxt))
            {
                textTxt.Text = edit.TextColour;
            }

            // Update swatches
            UpdateSwatch($"fill_{edit.StyleKey}", ParseColourOrDefault(edit.FillColour));
            UpdateSwatch($"stroke_{edit.StyleKey}", ParseColourOrDefault(edit.StrokeColour));
            UpdateSwatch($"text_{edit.StyleKey}", ParseColourOrDefault(edit.TextColour));
        }

        // Metrics
        foreach (var kvp in request.MetricEdits)
        {
            if (_metricNumericUpDowns.TryGetValue(kvp.Key, out var nud))
            {
                nud.Value = (decimal)kvp.Value;
            }
        }
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        // Validate all colour text boxes
        foreach (var kvp in _colorTextBoxes)
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

        Dictionary<string, double> metricEdits = [];
        foreach (var kvp in _metricNumericUpDowns)
        {
            metricEdits[kvp.Key] = (double)kvp.Value.Value;
        }

        // Build request
        List<ThemeStyleEdit> styleEdits = [];
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            string fill = _colorTextBoxes.TryGetValue($"fill_{preset.StyleKey}", out var ft) ? ft.Text : "";
            string stroke = _colorTextBoxes.TryGetValue($"stroke_{preset.StyleKey}", out var st) ? st.Text : "";
            string text = _colorTextBoxes.TryGetValue($"text_{preset.StyleKey}", out var tt) ? tt.Text : "";

            if (fill.Length > 0 || stroke.Length > 0 || text.Length > 0)
            {
                styleEdits.Add(new ThemeStyleEdit(preset.StyleKey, fill, stroke, text));
            }
        }

        var request = new ThemeEditRequest(styleEdits, metricEdits);

        if (!ThemeEditRequest.TryValidate(request, out _, out string? msg))
        {
            _warningLabel.Text = msg ?? "Invalid theme edit.";
            _warningLabel.Visible = true;
            DialogResult = DialogResult.None;
            return;
        }

        var warnings = ThemeEditRequest.ContrastWarnings(request);
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
        var confirm = MessageBox.Show(
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
        foreach (var txt in _colorTextBoxes.Values)
        {
            txt.Text = string.Empty;
        }

        // Reset every editable metric to its catalogue default.
        foreach (var kvp in _metricNumericUpDowns)
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