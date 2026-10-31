using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Scene;

namespace GanttCreator.AddIn;

/// <summary>
/// Read-only panel separating last-rendered scene warnings from current
/// table-validation findings. Never mutates workbook data.
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
    private readonly Label _renderedAtLabel = new()
    {
        Dock = DockStyle.Top,
        AutoSize = true,
        Padding = new Padding(8),
    };

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
    /// <param name="renderedAtUtc">The timestamp of the last successful render, or null when none is known.</param>
    /// <param name="currentIssues">Current validation findings, separate from the last rendered scene warnings.</param>
    internal WarningsPanelForm(
        IReadOnlyList<SceneWarning>? warnings,
        DateTimeOffset? renderedAtUtc = null,
        IReadOnlyList<GanttValidationIssue>? currentIssues = null)
    {
        InitializeComponent();
        _renderedAtLabel.Text = renderedAtUtc is { } renderedAt
            ? $"Warnings from the last successful render: {renderedAt.ToLocalTime():g}"
            : "No successful chart render has been recorded in this session.";
        Populate(warnings ?? [], currentIssues ?? []);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _listView.Dispose();
            _renderedAtLabel.Dispose();
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

        _ = _listView.Columns.Add("Source", 125, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Row / owner", 100, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Code / field", 150, HorizontalAlignment.Left);
        _ = _listView.Columns.Add("Message", 300, HorizontalAlignment.Left);
        Controls.Add(_listView);
        Controls.Add(_renderedAtLabel);

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

    [SuppressMessage("Microsoft.Globalization", "CA1303:Do not pass literals as localized parameters",
        Justification = "Dialog UI strings are not localized in this release.")]
    private void Populate(IReadOnlyList<SceneWarning> warnings, IReadOnlyList<GanttValidationIssue> currentIssues)
    {
        _listView.Items.Clear();
        foreach (SceneWarning warning in warnings
            .OrderBy(item => item.OwnerId.Kind)
            .ThenBy(item => item.OwnerId.Value, StringComparer.Ordinal)
            .ThenBy(item => item.Code, StringComparer.Ordinal)
            .ThenBy(item => item.Message, StringComparer.Ordinal))
        {
            ListViewItem item = new()
            {
                Text = "Last rendered",
            };
            _ = item.SubItems.Add(warning.OwnerId.Kind == SceneOwnerKind.Row
                    ? warning.OwnerId.Value
                    : "Chart");
            _ = item.SubItems.Add(warning.Code);
            _ = item.SubItems.Add(warning.Message);
            _ = _listView.Items.Add(item);
        }

        foreach (GanttValidationIssue issue in currentIssues
            .OrderBy(item => item.RowNumber)
            .ThenBy(item => item.Field, StringComparer.Ordinal)
            .ThenBy(item => item.Code, StringComparer.Ordinal))
        {
            ListViewItem item = new("Current validation");
            _ = item.SubItems.Add(issue.RowNumber.ToString(CultureInfo.InvariantCulture));
            _ = item.SubItems.Add($"{issue.Severity} · {issue.Field} · {issue.Code}");
            _ = item.SubItems.Add(issue.Message);
            _ = _listView.Items.Add(item);
        }

        if (_listView.Items.Count == 0)
        {
            _ = _listView.Items.Add(new ListViewItem(["—", "—", "—", "No scene warnings or current validation findings."]));
        }
    }
}
