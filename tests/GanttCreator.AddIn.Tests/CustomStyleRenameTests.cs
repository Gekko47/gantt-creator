using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using GanttCreator.Core;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// Tests for the R5.6a custom-style rename (WI-H): the Core already had a
/// validated <c>Rename</c> operation and the Office writer already applied
/// it, but no path ever EMITTED one, so the roadmap's "create, rename, and
/// delete" named a capability the user could not reach.
/// </summary>
public sealed class CustomStyleRenameTests
{
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

    private static TextBox RequireRenameKeyText(ThemeSettingsForm form)
    {
        var text = typeof(ThemeSettingsForm)
            .GetField("_customRenameKeyText", BindingFlags.NonPublic | BindingFlags.Instance)
            ?.GetValue(form) as TextBox;
        return Assert.IsType<TextBox>(text);
    }

    private static TabPage GetCustomStylesTab(ThemeSettingsForm form)
    {
        var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
        foreach (TabPage page in tabControl.TabPages)
        {
            if (page.Text == "Custom styles")
            {
                return page;
            }
        }

        throw new InvalidOperationException("Tab not found: Custom styles");
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

    private static ThemeSettingsForm FormWithSavedStyle()
    {
        return new ThemeSettingsForm(new ThemeEditRequest(
            [],
            new Dictionary<string, double>(),
            null,
            new List<CustomStyleDef>
            {
                new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
            }));
    }

    [Fact]
    public void A_rename_is_emitted_as_a_single_rename_operation()
    {
        // The capability WI-H exists for: the roadmap's R5.6a row says
        // "create, rename, and delete", and until now only the first and the
        // last were reachable.
        RunOnStaThread(() =>
        {
            using var form = FormWithSavedStyle();
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            list.SelectedItem = "SAVED";
            InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

            RequireRenameKeyText(form).Text = "RENAMED";
            InvokeClick(form, "CustomRenameButton_Click");

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            var edit = Assert.Single(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
            Assert.Equal(CustomStyleOperation.Rename, edit.Operation);
            Assert.Equal("SAVED", edit.StyleKey);
            Assert.Equal("RENAMED", edit.NewStyleKey);
        });
    }

    [Fact]
    public void A_rename_carries_the_styles_properties()
    {
        // The rename must not lose the style's properties: it is a rename, not
        // a delete plus a create, so the row keeps its formatting.
        RunOnStaThread(() =>
        {
            using var form = FormWithSavedStyle();
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            list.SelectedItem = "SAVED";
            InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

            RequireRenameKeyText(form).Text = "RENAMED";
            InvokeClick(form, "CustomRenameButton_Click");

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            var edit = Assert.Single(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
            Assert.Equal("#FF0000", edit.FillColour);
            Assert.Equal("#00FF00", edit.StrokeColour);
            Assert.Equal("#0000FF", edit.TextColour);
        });
    }

    [Fact]
    public void A_rename_is_not_also_a_delete_or_a_create()
    {
        // A renamed style's OLD key still exists in the workbook -- under its
        // new name -- so a Delete for the old key would remove it, and a
        // Create for the new key would be a duplicate.
        RunOnStaThread(() =>
        {
            using var form = FormWithSavedStyle();
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            list.SelectedItem = "SAVED";
            InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

            RequireRenameKeyText(form).Text = "RENAMED";
            InvokeClick(form, "CustomRenameButton_Click");

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            var edits = request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>();
            Assert.Single(edits);
            Assert.DoesNotContain(edits, e => e.Operation == CustomStyleOperation.Delete);
            Assert.DoesNotContain(edits, e => e.Operation == CustomStyleOperation.Create);
        });
    }

    [Fact]
    public void A_rename_to_a_built_in_key_is_refused()
    {
        // The built-in protection R5.6a requires. The refusal happens in the
        // dialog, so the user is told before anything is written.
        RunOnStaThread(() =>
        {
            using var form = FormWithSavedStyle();
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            list.SelectedItem = "SAVED";
            InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

            RequireRenameKeyText(form).Text = "AsPlannedActivity";
            InvokeClick(form, "CustomRenameButton_Click");

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            Assert.Empty(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
        });
    }

    [Fact]
    public void A_rename_to_an_existing_key_is_refused()
    {
        RunOnStaThread(() =>
        {
            using var form = new ThemeSettingsForm(new ThemeEditRequest(
                [],
                new Dictionary<string, double>(),
                null,
                new List<CustomStyleDef>
                {
                    new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
                    new("OTHER", "OTHER", "#00FF00", "#FF0000", "#0000FF"),
                }));
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            list.SelectedItem = "SAVED";
            InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

            RequireRenameKeyText(form).Text = "OTHER";
            InvokeClick(form, "CustomRenameButton_Click");

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            Assert.Empty(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
        });
    }

    [Fact]
    public void An_untouched_style_emits_no_rename()
    {
        // The control: without it, a rename would be emitted for every style
        // whether the user asked for one or not.
        RunOnStaThread(() =>
        {
            using var form = FormWithSavedStyle();
            form.CreateControl();

            InvokeClick(form, "OkButton_Click");

            var request = form.GetValidatedRequest();
            Assert.NotNull(request);
            Assert.Empty(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
        });
    }

    [Fact]
    public void Creating_a_style_named_after_a_built_in_is_refused()
    {
        // The same dead check, on the CREATE path. The key field forces upper
        // case and every built-in preset key is mixed case, so an ORDINAL
        // comparison could never match: the built-in protection was dead code
        // on both paths, and a style whose key differed from a built-in's only
        // by case could be created.
        RunOnStaThread(() =>
        {
            using var form = new ThemeSettingsForm(new ThemeEditRequest(
                [],
                new Dictionary<string, double>()));
            form.CreateControl();

            var tab = GetCustomStylesTab(form);
            var list = Assert.IsType<ListBox>(Find(tab, c => c is ListBox));
            Assert.NotNull(list);

            // The new-key box is beside the New button.
            var newButton = Assert.IsType<Button>(Find(tab, c => c is Button b && b.Text == "New custom style"));
            var container = newButton.Parent ?? throw new InvalidOperationException("New button has no parent.");
            var keyBox = container.Controls.OfType<TextBox>().Single();
            keyBox.Text = "AsPlannedActivity";

            InvokeClick(form, "CustomNewButton_Click");

            // The style was NOT created: the list is still empty.
            Assert.Empty(list.Items);
        });
    }
}
