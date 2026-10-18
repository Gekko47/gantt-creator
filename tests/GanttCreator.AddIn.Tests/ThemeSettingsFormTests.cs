using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using GanttCreator.AddIn;
using GanttCreator.Core;
using Xunit;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Tests for the R5.6a "Custom styles" tab of <see cref="ThemeSettingsForm"/>.
/// <para>
/// The dialog is exercised on an STA thread so WinForms handles exist for
/// list-box selection and button focus. The existing test hook
/// <see cref="ThemeSettingsForm.SetValidatedRequestForTesting"/> is not used
/// here because these tests assert the real button handlers run.
/// </para>
/// </summary>
public class ThemeSettingsFormTests
{
    [Fact]
    public void EmptyState_NewButtonAlwaysEnabledEditGroupDisabled()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();

                var tab = GetCustomStylesTab(form);
                var newButton = RequireButton(tab, "New custom style");
                var deleteButton = RequireButton(tab, "Delete");
                var editGroup = RequireGroupBox(tab, "Edit selected style");
                var list = RequireListBox(tab);
                var hint = RequireHint(tab);

                Assert.True(newButton.Enabled, "New custom style must be clickable with no styles defined.");
                Assert.False(editGroup.Enabled, "Edit panel must be disabled with no selection.");
                Assert.False(deleteButton.Enabled, "Delete must be disabled with no selection.");
                Assert.Equal("No custom styles defined yet. Enter a key and click New custom style.", hint.Text);
                Assert.Empty(list.Items);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void SavedStyles_LoadAndSelectEnableEditPanel()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [], new Dictionary<string, double>(), null,
                new List<CustomStyleDef>
                {
                    new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
                }));
            try
            {
                form.CreateControl();

                var tab = GetCustomStylesTab(form);
                var list = RequireListBox(tab);
                Assert.True(list.Items.Count == 1);
                Assert.True(list.Items.Contains("SAVED"));

                // Selecting the loaded style (as a user would) enables the edit panel.
                list.SelectedIndex = 0;

                Assert.True(RequireGroupBox(tab, "Edit selected style").Enabled);
                var hint = RequireHint(tab);
                _ = hint.Handle; // force handle creation so Visible can be read
                Assert.False(hint.Visible, "Empty-state hint must be hidden when styles exist.");

                var edits = GetField<Dictionary<string, CustomStyleState>>(form, "_customStyleEdits");
                Assert.True(edits.ContainsKey("SAVED"));
                var state = edits["SAVED"];
                Assert.Equal("#FF0000", state.Fill);
                Assert.Equal("#00FF00", state.Stroke);
                Assert.Equal("#0000FF", state.Text);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void CreateNewStyle_Valid_AddsWithDefaultColoursAndEnablesEdit()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();

                var tab = GetCustomStylesTab(form);
                var keyText = RequireKeyTextBox(tab);
                keyText.Text = "NEWSTYLE";

                InvokeClick(form, "CustomNewButton_Click");

                var list = RequireListBox(tab);
                Assert.True(list.Items.Contains("NEWSTYLE"));
                Assert.Equal("NEWSTYLE", list.SelectedItem);

                Assert.True(RequireGroupBox(tab, "Edit selected style").Enabled);
                Assert.False(RequireHint(tab).Visible);
                Assert.Equal(string.Empty, keyText.Text);

                var edits = GetField<Dictionary<string, CustomStyleState>>(form, "_customStyleEdits");
                var entry = Assert.Single(edits);
                Assert.Equal("NEWSTYLE", entry.Key);
                Assert.Equal("#FFFFFF", entry.Value.Fill);
                Assert.Equal("#000000", entry.Value.Stroke);
                Assert.Equal("#000000", entry.Value.Text);

                var warning = GetField<Label>(form, "_warningLabel");
                Assert.False(warning.Visible);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void CreateNewStyle_BlankKey_RejectedWithWarning()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                var keyText = RequireKeyTextBox(tab);
                keyText.Text = "   ";

                InvokeClick(form, "CustomNewButton_Click");

                var warning = GetField<Label>(form, "_warningLabel");
                Assert.Contains("Enter a key", warning.Text, StringComparison.Ordinal);
                Assert.Empty(RequireListBox(tab).Items);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void CreateNewStyle_UppercaseKey_CreatedAsCustomStyle()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                RequireKeyTextBox(tab).Text = "MYKEY";

                InvokeClick(form, "CustomNewButton_Click");

                var warning = GetField<Label>(form, "_warningLabel");
                Assert.Empty(warning.Text);
                var list = RequireListBox(tab);
                Assert.True(list.Items.Count == 1);
                Assert.True(list.Items.Contains("MYKEY"));
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void CreateNewStyle_DuplicateKey_RejectedWithWarning()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [], new Dictionary<string, double>(), null,
                new List<CustomStyleDef> { new("DUP", "DUP", "#000000", "#FFFFFF", "#000000") }));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                RequireKeyTextBox(tab).Text = "DUP";

                InvokeClick(form, "CustomNewButton_Click");

                var warning = GetField<Label>(form, "_warningLabel");
                Assert.Contains("already exists", warning.Text, StringComparison.Ordinal);
                var list = RequireListBox(tab);
                Assert.True(list.Items.Count == 1);
                Assert.True(list.Items.Contains("DUP"));
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void Ok_UntouchedSavedStyle_EmitsNoCustomStyleEdits()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [], new Dictionary<string, double>(), null,
                new List<CustomStyleDef>
                {
                    new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
                }));
            try
            {
                form.CreateControl();

                // No edits made; opening the dialog must not schedule a delete
                // for a saved style that is still present.
                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                Assert.Empty(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
            }
            finally
            {
                form.Dispose();
            }
        });
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
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join(TimeSpan.FromSeconds(10));
        tcs.Task.GetAwaiter().GetResult();
    }

    private static TabPage GetCustomStylesTab(Form form)
    {
        var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
        foreach (TabPage page in tabControl.TabPages)
        {
            if (page.Text == "Custom styles")
            {
                return page;
            }
        }
        throw new InvalidOperationException("Custom styles tab not found.");
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

    private static Button RequireButton(TabPage tab, string text) =>
        Assert.IsType<Button>(Find(tab, c => c is Button b && b.Text == text));

    private static ListBox RequireListBox(TabPage tab) =>
        Assert.IsType<ListBox>(Find(tab, c => c is ListBox));

    private static GroupBox RequireGroupBox(TabPage tab, string text) =>
        Assert.IsType<GroupBox>(Find(tab, c => c is GroupBox g && g.Text == text));

    private static Label RequireHint(TabPage tab) =>
        Assert.IsType<Label>(Find(tab, c => c is Label l && l.Text.StartsWith("No custom styles", StringComparison.Ordinal)));

    private static TextBox RequireKeyTextBox(TabPage tab)
    {
        var newButton = RequireButton(tab, "New custom style");
        var container = newButton.Parent ?? throw new InvalidOperationException("New button has no parent.");
        foreach (Control c in container.Controls)
        {
            if (c is TextBox tb)
            {
                return tb;
            }
        }
        throw new InvalidOperationException("New key textbox not found.");
    }

    private static TField GetField<TField>(ThemeSettingsForm form, string name)
    {
        var f = typeof(ThemeSettingsForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException($"Field '{name}' not found on ThemeSettingsForm.");
        return (TField)f.GetValue(form)!;
    }

    private static void InvokeClick(ThemeSettingsForm form, string methodName)
    {
        var m = typeof(ThemeSettingsForm).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException($"Method '{methodName}' not found on ThemeSettingsForm.");
        m.Invoke(form, new object?[] { null, EventArgs.Empty });
    }
}
