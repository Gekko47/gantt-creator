using System.Diagnostics.CodeAnalysis;
using GanttCreator.Core.Scene;

namespace GanttCreator.AddIn;

/// <summary>
/// Read-only WinForms panel listing the scene-level warnings from the
/// last Refresh, ordered severity → row → kind. Never mutates data: the panel
/// only displays warnings produced by the scene builder.
/// </summary>
/// <remarks>
/// <para>
/// Dialog pattern follows <see cref="ChartTitleForm"/>: modal via
/// ShowDialog, runs on the Excel main STA thread, fully
/// offline.
/// </para>
/// </remarks>
internal sealed partial class WarningsPanelForm : Form
{
    private readonly ListView _listView = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        MultiSelect = false,
        HideSelection = false,
    };

    private readonly Button _closeButton = new()
    {
        Text = "&Close",
        DialogResult = DialogResult.OK,
        AutoSize = true,
        Margin = new Padding(4, 0, 0, 0),
    };

    /// <summary>
    /// Initialises the panel with the warnings from the last Refresh.
    /// </summary>
    /// <param name="warnings">The scene-level warnings, or null for none.</param>
    internal WarningsPanelForm(IReadOnlyList<SceneWarning>? warnings)
    {
        InitializeComponent();
        Populate(warnings ?? []);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _listView.Dispose();
            _closeButton.Dispose();
        }
        base.Dispose(disposing);
    }

    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void InitializeComponent()
    {
        Text = "Gantt Creator Warnings";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(600, 400);
        AutoScaleMode = AutoScaleMode.Dpi;

        _ = _listView.Columns.Add("Severity", 80, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Row", 50, HorizontalAlignment.Right);
        _ = _listView.Columns.Add("Field", 100, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Code", 150, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Message", 300, HorizontalAlignment.Left);
        Controls.Add(_listView);

        FlowLayoutPanel buttonPanel = new()
        {
            Dock = DockStyle.Bottom,
            Height = 40,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(6),
        };
        buttonPanel.Controls.Add(_closeButton);
        Controls.Add(buttonPanel);

        AcceptButton = _closeButton;
        CancelButton = _closeButton;
        _closeButton.Click += (_, _) => Close();
    }

    private void Populate(IReadOnlyList<SceneWarning> warnings)
    {
        _listView.Items.Clear();
        foreach (SceneWarning warning in warnings)
        {
            ListViewItem item = new()
            {
                Text = warning.OwnerId.Kind == SceneOwnerKind.Row
                    ? warning.OwnerId.Value
                    : "Chart",
            };
            _ = item.SubItems.Add(warning.Code);
            _ = item.SubItems.Add(warning.Message);
            _ = _listView.Items.Add(item);
        }

        if (_listView.Items.Count == 0)
        {
            _ = _listView.Items.Add(new ListViewItem(["—", "—", "—", "—", "No warnings."]));
        }
    }
}
