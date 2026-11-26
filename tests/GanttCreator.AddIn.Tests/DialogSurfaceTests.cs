using System.Reflection;
using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The three dialogs that had no test constructing them (audit finding 5).
/// </summary>
/// <remarks>
/// Their decision logic is already tested in Core (<c>ChartTitleEditRequestTests</c>), so these
/// facts pin the dialog surface itself: the accept handler produces the validated request or stays
/// open, the picker applies the bounds the service passes it, and the panel lists the findings it
/// was handed. The handler is invoked by reflection because it is private and the form is never
/// shown; <c>CustomStyleRenameTests</c> already drives the theme dialog the same way.
/// </remarks>
public class DialogSurfaceTests
{
    [Fact]
    public void The_title_dialog_returns_the_validated_request_when_accepted()
    {
        using ChartTitleForm form = new("As-Built", showTitle: true);
        form.CreateControl();

        Accept(form);

        ChartTitleEditRequest? request = form.GetValidatedRequest();
        Assert.NotNull(request);
        Assert.Equal("As-Built", request.Title);
        Assert.True(request.ShowTitle);
        Assert.Equal(DialogResult.OK, form.DialogResult);
    }

    [Fact]
    public void The_title_dialog_stays_open_for_a_blank_title()
    {
        using ChartTitleForm form = new("   ", showTitle: true);
        form.CreateControl();

        Accept(form);

        Assert.Null(form.GetValidatedRequest());
        Assert.Equal(DialogResult.None, form.DialogResult);
    }

    [Fact]
    public void The_date_picker_applies_the_lower_bound_it_was_given()
    {
        DateOnly initial = new(2026, 5, 20);
        DateOnly minimum = new(2026, 5, 1);

        using PlotDatePickerForm form = new("Plot finish", initial, minimum, max: null);

        MonthCalendar calendar = Find<MonthCalendar>(form);
        Assert.Equal(minimum.ToDateTime(TimeOnly.MinValue), calendar.MinDate);
        Assert.Equal(initial.ToDateTime(TimeOnly.MinValue), calendar.SelectionStart);
        Assert.Null(form.SelectedDate);
    }

    [Fact]
    public void The_warnings_panel_lists_the_findings_it_was_given()
    {
        GanttValidationIssue issue = new(
            RowNumber: 4,
            Field: "Finish",
            Code: "FinishBeforeStart",
            Severity: GanttValidationSeverity.Error,
            Message: "Finish is before Start.");

        using WarningsPanelForm form = new(warnings: null, renderedAtUtc: null, currentIssues: [issue]);

        ListView list = Find<ListView>(form);
        Assert.Contains(
            list.Items.Cast<ListViewItem>(),
            item => item.SubItems.Cast<ListViewItem.ListViewSubItem>()
                .Any(sub => sub.Text.Contains("FinishBeforeStart", StringComparison.Ordinal)));
    }

    /// <summary>Invokes the title dialog's private accept handler, as a click would.</summary>
    /// <param name="form">The dialog under test.</param>
    private static void Accept(ChartTitleForm form)
    {
        MethodInfo? handler = typeof(ChartTitleForm).GetMethod(
            "OkButton_Click",
            BindingFlags.NonPublic | BindingFlags.Instance);
        Assert.NotNull(handler);
        _ = handler.Invoke(form, [null, EventArgs.Empty]);
    }

    /// <summary>Finds the first control of the given type anywhere in the dialog.</summary>
    /// <typeparam name="T">The control type to find.</typeparam>
    /// <param name="root">The dialog to search.</param>
    /// <returns>The control.</returns>
    /// <exception cref="InvalidOperationException">The dialog holds no such control.</exception>
    private static T Find<T>(Control root)
        where T : Control =>
        FindOrDefault<T>(root)
        ?? throw new InvalidOperationException($"The dialog holds no {typeof(T).Name}.");

    private static T? FindOrDefault<T>(Control root)
        where T : Control
    {
        foreach (Control child in root.Controls)
        {
            if (child is T match)
            {
                return match;
            }

            if (FindOrDefault<T>(child) is { } nested)
            {
                return nested;
            }
        }

        return null;
    }
}
