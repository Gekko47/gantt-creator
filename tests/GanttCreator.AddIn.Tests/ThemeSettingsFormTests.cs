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
    public void ChartTab_OffersOneRowPerChartChromeRole()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();

                var chartTab = GetTab(form, "Chart");
                var appearance = RequireGroupBox(chartTab, "Chart appearance");

                var textBoxes = appearance.Controls.OfType<TextBox>().ToArray();
                Assert.Equal(GanttChartChrome.Roles.Count, textBoxes.Length);

                // Every role's row shows the approved default, and the rows are
                // the roles the contract names.
                foreach (var role in GanttChartChrome.Roles)
                {
                    var textBox = textBoxes.Single(
                        box => box.Name == GanttChartChrome.SettingKeyFor(role));
                    Assert.Equal(GanttChartChrome.DefaultHexFor(role), textBox.Text);
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void ChartTab_AStoredChromeColourIsShownAndCarriedIntoTheRequest()
    {
        RunOnStaThread(() =>
        {
            var request = new ThemeEditRequest(
                [],
                new Dictionary<string, double>(),
                ChromeColourEdits: new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["ChartBackgroundColour"] = "#F0F0F0",
                });
            var form = new ThemeSettingsForm(request);
            try
            {
                form.CreateControl();

                var chartTab = GetTab(form, "Chart");
                var appearance = RequireGroupBox(chartTab, "Chart appearance");
                var textBox = appearance.Controls
                    .OfType<TextBox>()
                    .Single(box => box.Name == "ChartBackgroundColour");

                // The dialog shows the workbook's colour, not the catalogue
                // default, so the user sees what the chart actually renders.
                Assert.Equal("#F0F0F0", textBox.Text);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

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

    [Fact]
    public void Ok_ColourChangeOnSavedCustomStyle_EmitsUpdateNotRename()
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

                // Select the saved custom style: loads its colours into the editor.
                var list = RequireListBox(tab);
                list.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // Change only the fill colour; the key stays "SAVED".
                GetField<TextBox>(form, "_customStyleFillText").Text = "#00FF00";
                InvokeClick(form, "CustomStyleColourTextChanged");

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edits = request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>();
                Assert.Single(edits);
                var edit = Assert.Single(edits);
                // A colour-only change is an upsert by key, not a rename into
                // an existing key, so it cannot trigger the duplicate-key conflict.
                Assert.Equal(CustomStyleOperation.Update, edit.Operation);
                Assert.Equal("SAVED", edit.StyleKey);
                Assert.Null(edit.NewStyleKey);
                Assert.Equal("#00FF00", edit.FillColour);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void CustomStyleEditor_OffersTheFillPatternAndTheStyleSizes()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);

                // The fill type, the outline width, the two sizes, and the
                // label vocabulary are properties of the STYLE, so they live in
                // the style editor rather than in the chart-wide metric tabs.
                var pattern = RequireComboBox(tab);
                Assert.Contains(
                    pattern.Items.Cast<object>().Select(item => item.ToString()),
                    label => label == "Solid");

                var numerics = FindAll(tab, c => c is NumericUpDown)
                    .Cast<NumericUpDown>()
                    .ToList();
                Assert.Equal(3, numerics.Count);

                var defaultLabel = FindAll(tab, c => c is ComboBox)
                    .Cast<ComboBox>()
                    .First(combo => combo.Items.Count > 0 && combo.Items[0] is not null);
                Assert.NotEmpty(defaultLabel.Items);

                var allowedLabels = Assert.IsType<CheckedListBox>(Find(tab, c => c is CheckedListBox));
                Assert.NotEmpty(allowedLabels.Items);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void Ok_PatternChangeOnSavedCustomStyle_CarriesThePattern()
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
                list.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // Choose a hatch pattern: the fill type is the property the
                // user asked for, and it must reach the emitted edit.
                var pattern = FindAll(tab, c => c is ComboBox)
                    .Cast<ComboBox>()
                    .First(combo => combo.Items.Cast<object>().Any(item => item.ToString() == "Cross"));
                pattern.SelectedItem = pattern.Items.Cast<object>()
                    .First(item => item.ToString() == "Cross");
                InvokeClick(form, "CustomStylePropertyChanged");

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edit = Assert.Single(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
                Assert.Equal(CustomStyleOperation.Update, edit.Operation);
                Assert.Equal(GanttHatchPattern.Cross, edit.HatchPattern);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_dialog_separates_entity_styles_chart_and_custom_styles()
    {
        // The three scopes are separate tabs, so the user can see which
        // options belong to a style object, which belong to the chart, and
        // which belong to a style they created.
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
                var titles = tabControl.TabPages.Cast<TabPage>().Select(page => page.Text).ToList();

                Assert.Equal(3, titles.Count);
                Assert.Equal("Entity styles", titles[0]);
                Assert.Equal("Chart", titles[1]);
                Assert.Equal("Custom styles", titles[2]);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_chart_tab_groups_the_options_by_the_object_they_change()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Chart");

                var groups = FindAll(tab, c => c is GroupBox)
                    .Cast<GroupBox>()
                    .Select(group => group.Text)
                    .ToList();

                // Rows and lanes, entity defaults, and grid are different
                // objects on the chart; the hatch spacing is a property of the
                // fill pattern rather than of the chart's geometry.
                Assert.Contains("Rows and lanes", groups);
                Assert.Contains("Entity defaults", groups);
                Assert.Contains("Grid", groups);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_entity_styles_tab_groups_the_built_in_styles_by_family()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                var familyLabels = FindAll(tab, c => c is Label)
                    .Cast<Label>()
                    .Select(label => label.Text)
                    .ToList();

                Assert.Contains("Activities", familyLabels);
                Assert.Contains("Milestones", familyLabels);
                Assert.Contains("Procurement", familyLabels);
                Assert.Contains("Chart elements", familyLabels);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void Ok_DashChangeOnSavedCustomStyle_CarriesTheDash()
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
                list.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // The stroke type is a first-class property: the line style
                // picker must reach the emitted edit.
                var dash = FindAll(tab, c => c is ComboBox)
                    .Cast<ComboBox>()
                    .First(combo => combo.Items.Cast<object>().Any(item => item.ToString() == "Dash"));
                dash.SelectedItem = dash.Items.Cast<object>().First(item => item.ToString() == "Dash");
                InvokeClick(form, "CustomStylePropertyChanged");

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edit = Assert.Single(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
                Assert.Equal(CustomStyleOperation.Update, edit.Operation);
                Assert.Equal(GanttStrokeDash.Dash, edit.StrokeDash);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void Ok_UntouchedSavedStyle_EmitsNoEdit()
    {
        // A style whose properties the user never touched must not produce an
        // edit: the stored pattern, sizes, and label vocabulary are carried
        // by the load, not by an edit.
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [], new Dictionary<string, double>(), null,
                new List<CustomStyleDef>
                {
                    new(
                        "SAVED",
                        "SAVED",
                        "#FF0000",
                        "#00FF00",
                        "#0000FF",
                        GanttHatchPattern.Cross,
                        0.75,
                        8.0,
                        12.0,
                        GanttLabelPosition.Inside,
                        new HashSet<GanttLabelPosition>
                        {
                            GanttLabelPosition.Auto,
                            GanttLabelPosition.Inside,
                        }),
                }));
            try
            {
                form.CreateControl();
                var list = RequireListBox(GetCustomStylesTab(form));
                list.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

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

    private static TabPage GetCustomStylesTab(Form form) => GetTab(form, "Custom styles");

    private static TabPage GetTab(Form form, string title)
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

    private static IEnumerable<Control> FindAll(Control root, Func<Control, bool> match)
    {
        if (match(root))
        {
            yield return root;
        }

        foreach (Control child in root.Controls)
        {
            foreach (var found in FindAll(child, match))
            {
                yield return found;
            }
        }
    }

    private static Button RequireButton(TabPage tab, string text) =>
        Assert.IsType<Button>(Find(tab, c => c is Button b && b.Text == text));

    private static ListBox RequireListBox(TabPage tab) =>
        Assert.IsType<ListBox>(Find(tab, c => c is ListBox));

    private static ComboBox RequireComboBox(TabPage tab) =>
        Assert.IsType<ComboBox>(Find(tab, c => c is ComboBox));

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
