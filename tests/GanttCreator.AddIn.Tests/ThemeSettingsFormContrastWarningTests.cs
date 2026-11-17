using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Tests for the R5.6 theme dialog's contrast warning: WI-C (the warning
/// actually reaches the user) and WI-D (an unchanged pair is not reported).
/// </summary>
public sealed class ThemeSettingsFormContrastWarningTests
{
    /// <summary>
    /// The Delay Event's shipped pair: white on red at about 4.0:1, which is
    /// below the 4.5:1 gate. It is the grandfathering case -- the shipped
    /// default is not the user's doing.
    /// </summary>
    private static ThemeEditRequest RequestWithDelayEvent() =>
        new(
            [new ThemeStyleEdit("DelayEvent", "#FF0000", "#C00000", "#FFFFFF")],
            new Dictionary<string, double>());

    /// <summary>
    /// Selects the Delay Event in the Entity styles list and shows its
    /// properties in the detail panel, so a colour the user changes is
    /// captured against that style.
    /// </summary>
    private static void SelectDelayEvent(ThemeSettingsForm form)
    {
        // The list view's selection is native state, which a unit test cannot
        // create, so the form's own seam shows the style instead.
        var seam = typeof(ThemeSettingsForm).GetMethod(
            "ShowEntityStyleForTesting",
            BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("ShowEntityStyleForTesting not found.");
        seam.Invoke(form, new object[] { "DelayEvent" });
    }

    /// <summary>
    /// Changes the selected style's FILL colour to the given value, the way
    /// the user does, and captures it.
    /// </summary>
    private static void SetSelectedFill(ThemeSettingsForm form, string hex)
    {
        var fill = typeof(ThemeSettingsForm)
            .GetField("_entityFillText", BindingFlags.NonPublic | BindingFlags.Instance)
            ?.GetValue(form) as TextBox;
        Assert.IsType<TextBox>(fill);
        fill.Text = hex;
    }

    [SuppressMessage("Design", "CA1031: Do not catch general exception types",
        Justification = "Any STA-thread exception must be marshaled back to the test thread and is rethrown via the task.")]
    private static void RunOnStaThread(Action action)
    {
        var tcs = new TaskCompletionSource();
        var thread = new Thread(() =>
        {
            try
            {
                action();
                tcs.SetResult();
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        })
        { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join(TimeSpan.FromSeconds(10));
        tcs.Task.GetAwaiter().GetResult();
    }

    private static void InvokeClick(ThemeSettingsForm form, string methodName)
    {
        var m = typeof(ThemeSettingsForm).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException($"Method '{methodName}' not found on ThemeSettingsForm.");
        m.Invoke(form, new object?[] { null, EventArgs.Empty });
    }

    private static Label RequireWarningLabel(ThemeSettingsForm form)
    {
        var label = typeof(ThemeSettingsForm)
            .GetField("_warningLabel", BindingFlags.NonPublic | BindingFlags.Instance)
            ?.GetValue(form) as Label;
        return Assert.IsType<Label>(label);
    }

    /// <summary>
    /// WI-C: the contrast warning is ADVISORY -- it never blocks the write --
    /// but it must be VISIBLE, and the dialog must STAY OPEN so the user
    /// reads it against the colours they just chose.
    /// </summary>
    /// <remarks>
    /// The warning used to be computed and then thrown away by the
    /// <c>DialogResult = OK</c> on the next line, so the dialog closed with
    /// the message set but never on screen. This is the positive test for
    /// that defect.
    /// </remarks>
    [Fact]
    public void A_low_contrast_edit_warns_visibly_and_still_writes()
    {
        RunOnStaThread(() =>
        {
            bool written = false;
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ =>
                {
                    written = true;
                    return true;
                });
            form.CreateControl();

            // The user changes the Delay Event's fill to white, leaving white
            // text on white: a violation the user created.
            SelectDelayEvent(form);
            SetSelectedFill(form, "#FFFFFF");

            InvokeClick(form, "OkButton_Click");

            // The warning is on screen. The label's effective Visible is
            // false in a test because the form itself was never shown -- a
            // control is visible only when its whole parent chain is -- so
            // the warning is asserted by its TEXT and by the dialog staying
            // open, which is what the user sees.
            var warning = RequireWarningLabel(form);
            Assert.NotEmpty(warning.Text);
            Assert.Contains("contrast", warning.Text, StringComparison.OrdinalIgnoreCase);

            // The write happened anyway: the warning is advisory.
            Assert.True(written, "The contrast warning blocked the write.");

            // The dialog stayed open, so the user can read the warning.
            Assert.Equal(DialogResult.None, form.DialogResult);
        });
    }

    /// <summary>
    /// WI-C: the SECOND OK is the acknowledgement. The same warning is not
    /// shown twice, so the dialog closes. Without this the dialog would be
    /// impossible to close.
    /// </summary>
    [Fact]
    public void The_second_ok_acknowledges_the_warning_and_closes()
    {
        RunOnStaThread(() =>
        {
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ => true);
            form.CreateControl();

            SelectDelayEvent(form);
            SetSelectedFill(form, "#FFFFFF");

            InvokeClick(form, "OkButton_Click");
            Assert.Equal(DialogResult.None, form.DialogResult);

            InvokeClick(form, "OkButton_Click");
            Assert.Equal(DialogResult.OK, form.DialogResult);
        });
    }

    /// <summary>
    /// WI-C: the write happens EXACTLY ONCE across the warning and its
    /// acknowledgement. The first OK applied the change; the second is the
    /// user saying "I have seen it". Writing twice would apply the same
    /// request to the workbook twice.
    /// </summary>
    [Fact]
    public void The_warning_and_its_acknowledgement_write_only_once()
    {
        RunOnStaThread(() =>
        {
            int writes = 0;
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ =>
                {
                    writes++;
                    return true;
                });
            form.CreateControl();

            SelectDelayEvent(form);
            SetSelectedFill(form, "#FFFFFF");

            InvokeClick(form, "OkButton_Click");
            Assert.Equal(1, writes);

            InvokeClick(form, "OkButton_Click");
            Assert.Equal(1, writes);
        });
    }

    /// <summary>
    /// WI-C: a clean edit shows no warning and closes the dialog at once.
    /// The warning must not turn every save into a two-step affair.
    /// </summary>
    [Fact]
    public void A_clean_edit_closes_without_a_warning()
    {
        RunOnStaThread(() =>
        {
            bool written = false;
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ =>
                {
                    written = true;
                    return true;
                });
            form.CreateControl();

            InvokeClick(form, "OkButton_Click");

            Assert.True(written, "The clean edit was not written.");
            Assert.Empty(RequireWarningLabel(form).Text);
            Assert.Equal(DialogResult.OK, form.DialogResult);
        });
    }

    /// <summary>
    /// WI-D: a stock workbook's shipped pair is GRANDFATHERED. The dialog
    /// emits an edit for every style it holds, not only the ones the user
    /// touched, so without the baseline the shipped Delay Event warned every
    /// time the dialog was opened and accepted -- about a colour the user
    /// never touched.
    /// </summary>
    [Fact]
    public void An_unchanged_shipped_pair_does_not_warn()
    {
        RunOnStaThread(() =>
        {
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ => true);
            form.CreateControl();

            InvokeClick(form, "OkButton_Click");

            Assert.Empty(RequireWarningLabel(form).Text);
            Assert.Equal(DialogResult.OK, form.DialogResult);
        });
    }

    /// <summary>
    /// WI-D: a user edit that CREATES a violating pair warns. This is the
    /// positive control for the grandfathering: without it, "no warning"
    /// would be indistinguishable from "the check never fires".
    /// </summary>
    [Fact]
    public void A_user_edit_that_creates_a_violation_warns()
    {
        RunOnStaThread(() =>
        {
            using var form = new ThemeSettingsForm(
                RequestWithDelayEvent(),
                _ => true);
            form.CreateControl();

            SelectDelayEvent(form);
            SetSelectedFill(form, "#FFFFFF");

            InvokeClick(form, "OkButton_Click");

            Assert.NotEmpty(
                RequireWarningLabel(form).Text);
            Assert.Equal(DialogResult.None, form.DialogResult);
        });
    }

    private static TabPage GetTab(ThemeSettingsForm form, string title)
    {
        var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
        foreach (TabPage page in tabControl.TabPages)
        {
            if (page.Text == title)
            {
                return page;
            }
        }

        throw new InvalidOperationException($"Tab not found: {title}");
    }

    private static Control? Find(Control root, Func<Control, bool> match)
    {
        if (match(root))
        {
            return root;
        }

        foreach (Control child in root.Controls)
        {
            var found = Find(child, match);
            if (found is not null)
            {
                return found;
            }
        }

        return null;
    }
}
