using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core;

namespace GanttCreator.AddIn;

/// <summary>
/// WinForms modal dialog for R5.6b chart-title settings (title text + visibility).
/// Runs on the Excel main STA thread via ShowDialog; fully offline.
/// </summary>
internal sealed partial class ChartTitleForm : Form
{
    private ChartTitleEditRequest? _validatedRequest;

    private readonly TextBox _titleTextBox = new()
    {
        Location = new Point(12, 30),
        Size = new Size(420, 23),
    };

    private readonly CheckBox _showTitleCheckBox = new()
    {
        Location = new Point(12, 62),
        Size = new Size(420, 24),
        Text = "Show title",
    };

    private readonly Button _okButton = new()
    {
        Text = "OK",
        DialogResult = DialogResult.OK,
        AutoSize = true,
    };

    private readonly Button _cancelButton = new()
    {
        Text = "Cancel",
        DialogResult = DialogResult.Cancel,
        AutoSize = true,
    };

    private readonly Label _warningLabel = new()
    {
        Dock = DockStyle.Top,
        Height = 24,
        ForeColor = Color.FromArgb(192, 0, 0),
        TextAlign = ContentAlignment.MiddleLeft,
        Visible = false,
    };

    /// <summary>
    /// Initializes the dialog with the current chart-title values.
    /// </summary>
    /// <param name="title">The current chart title.</param>
    /// <param name="showTitle">Whether the populated title band is visible.</param>
    internal ChartTitleForm(string title, bool showTitle)
    {
        InitializeComponent();
        _titleTextBox.Text = title;
        _showTitleCheckBox.Checked = showTitle;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _titleTextBox.Dispose();
            _showTitleCheckBox.Dispose();
            _warningLabel.Dispose();
            _cancelButton.Dispose();
            _okButton.Dispose();
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// Gets the validated request if the user clicked OK, or null if cancelled/invalid.
    /// </summary>
    internal ChartTitleEditRequest? GetValidatedRequest() => _validatedRequest;

    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void InitializeComponent()
    {
        Text = "Chart title";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(450, 150);
        AutoScaleMode = AutoScaleMode.Dpi;

        Label titleLabel = new Label
        {
            Text = "Title",
            Location = new Point(12, 10),
            AutoSize = true,
        };
        Controls.Add(titleLabel);
        Controls.Add(_titleTextBox);
        Controls.Add(_showTitleCheckBox);
        Controls.Add(_warningLabel);

        FlowLayoutPanel buttonPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 40,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(6),
        };
        buttonPanel.Controls.Add(_okButton);
        buttonPanel.Controls.Add(_cancelButton);
        Controls.Add(buttonPanel);

        AcceptButton = _okButton;
        CancelButton = _cancelButton;
        _okButton.Click += OkButton_Click;
    }

    private void OkButton_Click(object? sender, EventArgs e)
    {
        if (!ChartTitleEditRequest.TryCreate(
                _titleTextBox.Text, _showTitleCheckBox.Checked,
                out ChartTitleEditRequest? request, out _, out string? message)
            || request is null)
        {
            _warningLabel.Text = message ?? "Invalid chart title.";
            _warningLabel.Visible = true;
            _ = _titleTextBox.Focus();
            DialogResult = DialogResult.None;
            return;
        }

        _validatedRequest = request;
        DialogResult = DialogResult.OK;
    }
}
