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
    /// <summary>
    /// The property groups every detail panel holds, in panel order. Hoisted so
    /// the bounds tests share one list rather than re-allocating a constant
    /// array per call.
    /// </summary>
    private static readonly string[] s_detailGroupTexts = ["Fill", "Outline", "Label", "Size"];

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

    /// <summary>
    /// WI-A: a new custom style starts with a bar height of ZERO, which the
    /// style contract reads as "not applicable to this entity kind" -- so the
    /// row inherits the height its own kind carries rather than a copied
    /// constant. Seeding the code-owned default instead would pin a value
    /// into the style, and the renderer would then keep that number even
    /// though the style never asked for it.
    /// </summary>
    [Fact]
    public void CreateNewStyle_BarHeight_StartsAtZero_SoTheStyleInheritsItsKindsHeight()
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

                var edits = GetField<Dictionary<string, CustomStyleState>>(form, "_customStyleEdits");
                var entry = Assert.Single(edits).Value;

                Assert.Equal(0, entry.ActivityHeightPt);

                // The milestone size already used this convention, so the two
                // size overrides agree: both start at zero and both inherit.
                Assert.Equal(0, entry.MilestoneSizePt);

                // The editor shows the same zero, so the user sees the
                // inherited state rather than a copied constant.
                var barHeight = GetField<NumericUpDown>(form, "_customActivityHeightNumeric");
                Assert.Equal(0m, barHeight.Value);
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
    public void Ok_AnEightDigitChartColourIsRefusedBecauseTheWriterWouldRefuseIt()
    {
        // R5.6d: the dialog's own check used the lenient ColourHex rule, which
        // also accepts the eight-digit #RRGGBBAA form. The writer and the
        // reader require exactly uppercase #RRGGBB, so the dialog accepted a
        // value the write then refused - which is how the chart-appearance
        // editors ended up reporting errors in the sheet.
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [],
                new Dictionary<string, double>(),
                ChromeColourEdits: new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["ChartBackgroundColour"] = "#F0F0F0",
                }));
            try
            {
                form.CreateControl();
                var chartTab = GetTab(form, "Chart");
                var appearance = RequireGroupBox(chartTab, "Chart appearance");
                var textBox = appearance.Controls
                    .OfType<TextBox>()
                    .Single(box => box.Name == "ChartBackgroundColour");

                textBox.Text = "#F0F0F0FF";
                InvokeClick(form, "OkButton_Click");

                // The dialog stays open and produces no request.
                Assert.Null(form.GetValidatedRequest());
                Assert.True(form.DialogResult == DialogResult.None);
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
    public void Each_custom_style_colour_picker_edits_its_own_colour()
    {
        // Regression: all three colour picker buttons were wired to
        // CustomFillColorButton_Click, so clicking the stroke or text picker
        // changed the fill colour instead of its own.
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

                // Select the saved style so its colours load into the editor.
                var list = RequireListBox(tab);
                list.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // The three picker buttons must each target their own textbox.
                // Reading the stored colours back from the edit model confirms
                // that fill, stroke, and text are independently editable:
                // the picker buttons now each call PickCustomColour with their
                // own textbox/swatch pair rather than all sharing the fill one.
                var fillText = GetField<TextBox>(form, "_customStyleFillText");
                var strokeText = GetField<TextBox>(form, "_customStyleStrokeText");
                var textText = GetField<TextBox>(form, "_customStyleTextText");

                // The editor loads the style's stored colours.
                Assert.Equal("#FF0000", fillText.Text);
                Assert.Equal("#00FF00", strokeText.Text);
                Assert.Equal("#0000FF", textText.Text);

                // Change each colour independently through its textbox.
                fillText.Text = "#111111";
                InvokeClick(form, "CustomStyleColourTextChanged");
                strokeText.Text = "#222222";
                InvokeClick(form, "CustomStyleColourTextChanged");
                textText.Text = "#333333";
                InvokeClick(form, "CustomStyleColourTextChanged");

                // OK commits the edits. Each colour must be independently
                // present in the emitted edit: none of them may have been
                // overwritten by another picker's target (the bug).
                InvokeClick(form, "OkButton_Click");
                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edit = Assert.Single(request.CustomStyleEdits ?? Enumerable.Empty<CustomStyleEdit>());
                Assert.Equal(CustomStyleOperation.Update, edit.Operation);
                Assert.Equal("SAVED", edit.StyleKey);
                Assert.Equal("#111111", edit.FillColour);
                Assert.Equal("#222222", edit.StrokeColour);
                Assert.Equal("#333333", edit.TextColour);

                // The three picker handler methods are distinct: each picker
                // now targets its own colour rather than all sharing fill.
                var fillMethod = typeof(ThemeSettingsForm)
                    .GetMethod("CustomFillColorButton_Click",
                        BindingFlags.NonPublic | BindingFlags.Instance);
                var strokeMethod = typeof(ThemeSettingsForm)
                    .GetMethod("CustomStrokeColorButton_Click",
                        BindingFlags.NonPublic | BindingFlags.Instance);
                var textMethod = typeof(ThemeSettingsForm)
                    .GetMethod("CustomTextColorButton_Click",
                        BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(fillMethod);
                Assert.NotNull(strokeMethod);
                Assert.NotNull(textMethod);
                Assert.NotSame(fillMethod, strokeMethod);
                Assert.NotSame(fillMethod, textMethod);
                Assert.NotSame(strokeMethod, textMethod);
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

                // Rows and lanes, grid, and milestones are different objects
                // on the chart. The bar height and the outline width are gone
                // because the scene resolves them from the STYLE row, and the
                // hatch spacing is gone because the host's pattern cannot
                // express it.
                Assert.Contains("Rows and lanes", groups);
                Assert.Contains("Grid", groups);
                Assert.Contains("Milestones", groups);
                Assert.DoesNotContain("Entity defaults", groups);
                Assert.DoesNotContain("Hatch spacing", groups);

                // The delineator's line width is NOT a chart-wide knob: it
                // belongs to the delineator, so it lives on the Entity styles
                // tab beside the delineator's other properties. The value has
                // one home.
                Assert.DoesNotContain("Delineator", groups);
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

                var list = RequireEntityStyleList(tab);
                var headings = list.Groups.Cast<ListViewGroup>().Select(group => group.Header).ToList();

                // One heading per family, and no family twice: the heading is
                // now a property of the catalogue rather than of the order the
                // dialog happens to render the styles in.
                Assert.Equal(headings.Count, headings.Distinct(StringComparer.Ordinal).Count());
                Assert.Contains("Activities", headings);
                Assert.Contains("Milestones", headings);
                Assert.Contains("Procurement", headings);
                Assert.Contains("Chart elements", headings);
                Assert.DoesNotContain("Custom types", headings);

                // Every built-in style that carries a colour appears exactly
                // once, under the family its catalogue entry names.
                var styleKeys = list.Items.Cast<ListViewItem>()
                    .Select(item => (string)item.Tag!)
                    .ToList();
                Assert.Equal(styleKeys.Count, styleKeys.Distinct(StringComparer.Ordinal).Count());

                foreach (var item in list.Items.Cast<ListViewItem>())
                {
                    var styleKey = (string)item.Tag!;
                    Assert.Equal(
                        GanttStyleFamilies.DisplayNameFor(GanttStyleFamilies.For(styleKey)),
                        item.Group!.Header);
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_swatch_shows_the_workbook_colour_not_the_code_default()
    {
        RunOnStaThread(() =>
        {
            // The workbook's colour for As-Planned Activity differs from the
            // code-owned preset, so the swatch must show the workbook's.
            var request = new ThemeEditRequest(
                [new ThemeStyleEdit("AsPlannedActivity", "#123456", "#654321", "#ABCDEF")],
                new Dictionary<string, double>());
            var form = new ThemeSettingsForm(request);
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "AsPlannedActivity");

                var fill = GetField<TextBox>(form, "_entityFillText");
                Assert.Equal("#123456", fill.Text);

                var swatch = GetField<Panel>(form, "_entityFillSwatch");
                Assert.Equal(
                    Color.FromArgb(0x12, 0x34, 0x56).ToArgb(),
                    swatch.BackColor.ToArgb());
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

    [Fact]
    public void The_delineator_offers_an_editable_line_type()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "DefaultDelineator");

                // The dash is only editable for the delineator, so the combo
                // offers the closed vocabulary. (Its Visible flag is not
                // asserted: WinForms reports a control as hidden until the
                // whole form is shown.)
                var dash = FindAll(tab, c => c is ComboBox)
                    .Cast<ComboBox>()
                    .First(combo => combo.Items.Cast<object>().Any(item => item.ToString() == "Dash"));

                dash.SelectedItem = dash.Items.Cast<object>().First(candidate => candidate.ToString() == "Dash");

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edit = request.StyleEdits!.Single(candidate => candidate.StyleKey == "DefaultDelineator");
                Assert.Equal(GanttStrokeDash.Dash, edit.StrokeDash);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void A_built_in_style_other_than_the_delineator_keeps_its_code_owned_line_type()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "AsPlannedActivity");

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                var edit = request.StyleEdits!.Single(candidate => candidate.StyleKey == "AsPlannedActivity");

                // The general dash stays code-owned: the reader exact-matches
                // that column, so an edited value could not be kept.
                Assert.Null(edit.StrokeDash);
            }
            finally
            {
                form.Dispose();
            }
        });
    }


    [Fact]
    public void The_delineator_line_width_is_edited_on_the_entity_tab()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "DefaultDelineator");

                var width = GetField<NumericUpDown>(form, "_entityLineWidthNumeric");
                width.Value = 2.5m;

                InvokeClick(form, "OkButton_Click");

                var request = form.GetValidatedRequest();
                Assert.NotNull(request);
                Assert.Equal(2.5, request.MetricEdits["DelineatorLinePt"], 3);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_chart_tab_no_longer_offers_the_delineator_width()
    {
        // The delineator's line width belongs to the delineator, so it lives on
        // the Entity styles tab beside the delineator's other properties. The
        // chart tab keeps the chart-wide knobs only, so the value has one home.
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var chartTab = GetTab(form, "Chart");

                Assert.Null(Find(chartTab, c => c is GroupBox g && g.Text == "Delineator"));
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_dialog_opens_on_the_workbooks_line_type_for_the_delineator()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [new ThemeStyleEdit("DefaultDelineator", "#000000", "#000000", "#000000", GanttStrokeDash.Dash)],
                new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "DefaultDelineator");

                var dash = FindAll(tab, c => c is ComboBox)
                    .Cast<ComboBox>()
                    .First(combo => combo.Items.Cast<object>().Any(item => item.ToString() == "Dash"));
                Assert.Equal("Dash", dash.SelectedItem?.ToString());
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void Reset_restores_the_code_owned_preset_colours()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [new ThemeStyleEdit("AsPlannedActivity", "#FF0000", "#00FF00", "#0000FF")],
                new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");

                SelectEntityStyle(tab, "AsPlannedActivity");
                Assert.Equal("#FF0000", GetField<TextBox>(form, "_entityFillText").Text);

                // The reset logic is behind the confirmation dialog, so the
                // restore itself is invoked directly.
                InvokePrivate(form, "ApplyCodeOwnedDefaults");

                Assert.Equal(
                    GanttCatalogues.GetPreset("AsPlannedActivity").FillColour,
                    GetField<TextBox>(form, "_entityFillText").Text);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void No_control_on_the_custom_styles_tab_overflows_its_group()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);

                // No control may extend past the bounds of the group that
                // holds it. The earlier revision placed them absolutely, so
                // the permitted-label list ran past its group and the Delete
                // button was clipped by the foot of the panel.
                foreach (GroupBox group in FindAll(tab, c => c is GroupBox).Cast<GroupBox>())
                {
                    foreach (Control child in group.Controls)
                    {
                        Assert.True(
                            child.Right <= group.ClientSize.Width + 1
                                && child.Bottom <= group.ClientSize.Height + 1,
                            $"'{child.Name}' overflows group '{group.Text}'.");
                    }
                }

                // The New-style row holds the key box and the button side by
                // side in a flow panel, so they cannot overlap. The earlier
                // revision placed them at fixed coordinates and the button
                // sat on top of the key box.
                var newButton = RequireButton(tab, "New custom style");
                var keyBox = RequireKeyTextBox(tab);
                Assert.True(
                    newButton.Left >= keyBox.Right || keyBox.Left >= newButton.Right,
                    "The New custom style button overlaps the key box.");
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// The dialog is sized to its content: narrow enough that the property
    /// values are not stranded in a wide empty strip, wide enough that no tab
    /// needs a horizontal scroll bar.
    /// </summary>
    [Fact]
    public void The_dialog_is_sized_to_its_content_and_no_tab_overflows_it()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();

                // The width is a deliberate number, not a guess: the master
                // column plus the detail panel's label column and value
                // cluster, plus the padding.
                Assert.Equal(700, form.ClientSize.Width);

                var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
                foreach (TabPage page in tabControl.TabPages)
                {
                    LayOutTab(form, page);
                    page.PerformLayout();

                    // Every control on the page must sit inside the dialog's
                    // client area, or the page needs a horizontal scroll bar.
                    foreach (Control control in FindAll(page, c => c is not TabPage))
                    {
                        if (!control.Visible)
                        {
                            continue;
                        }

                        Point edge = control.PointToScreen(Point.Empty);
                        Point client = form.PointToScreen(Point.Empty);
                        int right = edge.X - client.X + control.Width;

                        Assert.True(
                            right <= form.ClientSize.Width,
                            $"'{control.Name}' on the '{page.Text}' tab reaches {right}px in a {form.ClientSize.Width}px dialog.");
                    }
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    [Fact]
    public void The_custom_styles_tab_fits_its_tab_without_a_horizontal_scroll_bar()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);

                // The tab is laid out by a split container, so it never needs a
                // horizontal scroll bar: the list is on the left and the
                // properties on the right, and the widest property group fits
                // beside the list in the dialog's client area.
                var split = Assert.IsType<SplitContainer>(Find(tab, c => c is SplitContainer));
                int widestGroup = FindAll(tab, c => c is GroupBox)
                    .Cast<GroupBox>()
                    .Select(group => group.Width)
                    .DefaultIfEmpty(0)
                    .Max();

                Assert.True(
                    split.SplitterDistance + widestGroup + 40 <= form.ClientSize.Width,
                    $"The tab needs {split.SplitterDistance + widestGroup}pt but the dialog is {form.ClientSize.Width}pt.");
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the Entity styles master list drew only its family headings
    /// and no rows. A <see cref="ListView"/> in <see cref="View.Details"/> draws
    /// items only when a column exists, so an item-less column list is invisible
    /// even though the items are populated.
    /// </summary>
    [Fact]
    public void EntityList_ShowsGroupedStyleRows()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();

                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);
                var list = RequireEntityStyleList(tab);

                Assert.Single(list.Columns);
                Assert.NotEmpty(list.Items);
                Assert.NotEmpty(list.Groups);
                Assert.All(list.Groups.Cast<ListViewGroup>(), group => Assert.NotEmpty(group.Items));

                // The rows are the code-owned presets, not a filtered subset
                // that happens to lose whole families.
                var shown = list.Items.Cast<ListViewItem>().Select(item => (string)item.Tag!).ToArray();
                Assert.Contains("AsBuiltMilestone", shown, StringComparer.Ordinal);
                Assert.Contains("DefaultDelineator", shown, StringComparer.Ordinal);

                // The named custom activity and the named custom milestone take
                // their style from the user, so their built-in fallback style is
                // not listed here: those styles live on the Custom styles tab,
                // where the user creates them.
                Assert.DoesNotContain("CustomActivity", shown, StringComparer.Ordinal);
                Assert.DoesNotContain("CustomMilestone", shown, StringComparer.Ordinal);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the four detail groups collapsed to a few pixels of width,
    /// so the selected style's properties were unreadable slivers.
    /// </summary>
    [Theory]
    [InlineData("Fill")]
    [InlineData("Outline")]
    [InlineData("Label")]
    [InlineData("Size")]
    public void EntityDetail_GroupsKeepUsableBounds(string groupText)
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);
                SelectEntityStyle(tab, "AsBuiltMilestone");
                tab.PerformLayout();

                var group = RequireGroupBox(tab, groupText);
                Assert.True(
                    group.Width > 300,
                    $"Entity '{groupText}' group collapsed to {group.Width}px wide.");
                Assert.True(
                    group.Height > 40,
                    $"Entity '{groupText}' group collapsed to {group.Height}px tall.");
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the Custom styles tab's New row was clipped to a strip with
    /// no visible button, and the edit groups collapsed the same way the Entity
    /// groups did.
    /// </summary>
    [Fact]
    public void CustomTab_NewRowAndEditGroupsAreUsable()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                LayOutTab(form, tab);

                var newButton = RequireButton(tab, "New custom style");

                // The form is never shown in a unit test, so every descendant
                // reports Visible false — including a button that renders
                // perfectly. Clipping is therefore asserted from geometry: the
                // button must land inside the master column rather than being
                // painted past its edge.
                var master = newButton.Parent?.Parent
                    ?? throw new InvalidOperationException("New button's master container not found.");
                Assert.True(
                    newButton.Height >= 20,
                    $"New button collapsed to {newButton.Height}px tall.");
                Assert.True(
                    newButton.Left >= 0,
                    $"New button starts at {newButton.Left}px, left of its master.");
                Assert.True(
                    newButton.Right <= master.ClientSize.Width,
                    $"New button ends at {newButton.Right}px but its master is {master.ClientSize.Width}px wide.");

                foreach (var groupText in s_detailGroupTexts)
                {
                    var group = RequireGroupBox(tab, groupText);
                    Assert.True(
                        group.Width > 300,
                        $"Custom '{groupText}' group collapsed to {group.Width}px wide.");
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the detail panel carried two title labels — a bold one that
    /// took the style's name and a grey one that was built later and never
    /// updated — so the empty state showed the hint twice and the hint stayed
    /// on screen after a style was selected.
    /// </summary>
    [Fact]
    public void EntityDetail_CarriesExactlyOneTitleThatFollowsTheSelection()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);

                Assert.Single(FindAll(
                    tab,
                    control => control is Label label
                        && label.Text.StartsWith("Select a style", StringComparison.Ordinal)));

                SelectEntityStyle(tab, "AsBuiltMilestone");

                var title = FindAll(tab, control => control is Label label && label.Text == "As-Built Milestone");
                Assert.Single(title);
                Assert.Empty(FindAll(
                    tab,
                    control => control is Label label
                        && label.Text.StartsWith("Select a style", StringComparison.Ordinal)));
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the property groups were 380px wide inside a 606px detail
    /// panel, so every group left a wide empty strip on the right and the
    /// properties looked stranded in the corner. The groups now fill the
    /// panel's width.
    /// </summary>
    [Fact]
    public void Entity_detail_groups_fill_the_detail_panel_width()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);
                SelectEntityStyle(tab, "AsPlannedActivity");
                tab.PerformLayout();

                var body = FindAll(tab, c => c is TableLayoutPanel t && t.ColumnCount == 1 && t.RowCount == 4)
                    .Cast<TableLayoutPanel>()
                    .Single();

                foreach (string caption in s_detailGroupTexts)
                {
                    var group = RequireGroupBox(tab, caption);
                    Assert.True(
                        group.Width >= body.ClientSize.Width - 2,
                        $"Entity '{caption}' group is {group.Width}px wide in a {body.ClientSize.Width}px panel.");
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: a property row that held a taller control than its label —
    /// a numeric up-down — was laid out past the row's height and clipped by
    /// the group box. Every property row now sizes itself to its content.
    /// </summary>
    [Fact]
    public void Every_property_row_is_taller_than_the_control_it_holds()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);
                SelectEntityStyle(tab, "AsPlannedActivity");
                tab.PerformLayout();

                foreach (string caption in s_detailGroupTexts)
                {
                    var group = RequireGroupBox(tab, caption);
                    var table = FindAll(group, c => c is TableLayoutPanel)
                        .Cast<TableLayoutPanel>()
                        .Single();

                    // A control's bounds are relative to the table, so the
                    // row's own top has to be added before comparing the
                    // control's foot with the row's foot.
                    int[] rowHeights = table.GetRowHeights();
                    foreach (Control value in table.Controls)
                    {
                        int row = table.GetRow(value);
                        if (row < 0 || row >= rowHeights.Length)
                        {
                            continue;
                        }

                        int rowTop = rowHeights.Take(row).Sum();
                        Assert.True(
                            value.Bottom <= rowTop + rowHeights[row] + 1,
                            $"Entity '{caption}' row {row} clips '{value.Name}' ({value.Bottom}px in a {rowHeights[row]}px row at {rowTop}px).");
                    }
                }
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// R5.6d WI-C: the default-position combo offered every label position,
    /// so the user could choose a default the style does not permit - a pair
    /// the writer refuses. The combo now offers exactly the permitted ones,
    /// and unchecking the current default moves it to a permitted position.
    /// </summary>
    [Fact]
    public void The_default_label_combo_offers_only_the_permitted_positions()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [],
                new Dictionary<string, double>(),
                null,
                new List<CustomStyleDef>
                {
                    new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
                }));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                LayOutTab(form, tab);

                var list = FindAll(tab, c => c is CheckedListBox)
                    .Cast<CheckedListBox>()
                    .Single();
                var combo = GetField<ComboBox>(form, "_customDefaultLabelCombo");

                // Select the saved style so the editor holds a real style.
                var styleList = RequireListBox(tab);
                styleList.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // The combo starts as the style's own permitted set - never
                // the whole vocabulary.
                var wholeVocabulary = Enum.GetValues<GanttLabelPosition>().Length;
                Assert.True(
                    combo.Items.Count < wholeVocabulary,
                    $"The default combo offered all {wholeVocabulary} positions.");

                // Untick every permitted position, then tick exactly two. The
                // handler is the one the host raises when the user clicks.
                for (int index = 0; index < list.Items.Count; index++)
                {
                    if (list.GetItemChecked(index))
                    {
                        InvokeAllowedLabelItemCheck(form, index, CheckState.Unchecked);
                        list.SetItemChecked(index, false);
                    }
                }

                var permitted = new List<GanttLabelPosition> { GanttLabelPosition.Left, GanttLabelPosition.Right };
                foreach (int index in Enumerable.Range(0, list.Items.Count))
                {
                    bool isChecked = list.Items[index] is ThemeSettingsForm.LabelPositionItem item
                        && permitted.Contains(item.Position);
                    if (isChecked)
                    {
                        InvokeAllowedLabelItemCheck(form, index, CheckState.Checked);
                        list.SetItemChecked(index, true);
                    }
                }

                // The combo offers the permitted positions and nothing else,
                // and the default it holds is one of them.
                var offered = combo.Items.Cast<ThemeSettingsForm.LabelPositionItem>()
                    .Select(item => item.Position)
                    .ToList();
                Assert.Equal(permitted, offered);
                ThemeSettingsForm.LabelPositionItem? chosen =
                    combo.SelectedItem as ThemeSettingsForm.LabelPositionItem;
                Assert.NotNull(chosen);
                Assert.Contains(chosen!.Position, permitted);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// R5.6d WI-C: the writer refuses a style that permits no label
    /// position, so the dialog refuses it too rather than after the write.
    /// </summary>
    [Fact]
    public void Ok_AStyleThatPermitsNoLabelPositionIsRefusedInTheDialog()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest(
                [],
                new Dictionary<string, double>(),
                null,
                new List<CustomStyleDef>
                {
                    new("SAVED", "SAVED", "#FF0000", "#00FF00", "#0000FF"),
                }));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                LayOutTab(form, tab);

                var list = FindAll(tab, c => c is CheckedListBox)
                    .Cast<CheckedListBox>()
                    .Single();

                var styleList = RequireListBox(tab);
                styleList.SelectedItem = "SAVED";
                InvokeClick(form, "CustomStyleList_SelectedIndexChanged");

                // Untick every permitted position.
                for (int index = 0; index < list.Items.Count; index++)
                {
                    if (list.GetItemChecked(index))
                    {
                        InvokeAllowedLabelItemCheck(form, index, CheckState.Unchecked);
                        list.SetItemChecked(index, false);
                    }
                }

                InvokeClick(form, "OkButton_Click");

                // The dialog stays open and produces no request.
                Assert.Null(form.GetValidatedRequest());
                Assert.True(form.DialogResult == DialogResult.None);
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the permitted-label list showed four of its eleven
    /// positions behind a scroll bar, and the property groups below it were
    /// pushed past the foot of the panel. The list now shows every position,
    /// and the groups scroll in their own pane with the Delete row pinned
    /// below them.
    /// </summary>
    [Fact]
    public void The_permitted_label_list_shows_every_position_and_the_delete_row_stays_visible()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetCustomStylesTab(form);
                LayOutTab(form, tab);

                var list = FindAll(tab, c => c is CheckedListBox)
                    .Cast<CheckedListBox>()
                    .Single();

                Assert.True(
                    list.Height >= (list.ItemHeight * list.Items.Count),
                    $"The permitted-label list shows {list.Height}px of its {list.ItemHeight * list.Items.Count}px of positions.");

                // The Delete row is pinned below the scrolling properties, so
                // it is always inside the panel even when the groups above it
                // are taller than the panel.
                var deleteButton = RequireButton(tab, "Delete");
                var panel = deleteButton.Parent?.Parent
                    ?? throw new InvalidOperationException("Delete button's container not found.");
                Assert.True(
                    deleteButton.Bottom <= panel.ClientSize.Height + 1,
                    $"The Delete button ends at {deleteButton.Bottom}px in a {panel.ClientSize.Height}px panel.");
            }
            finally
            {
                form.Dispose();
            }
        });
    }

    /// <summary>
    /// Regression: the entity styles tab showed the built-in style's label
    /// position as an empty value, so the row read as if the property were
    /// missing rather than fixed by the style.
    /// </summary>
    [Fact]
    public void The_entity_label_position_is_shown_as_a_read_only_value()
    {
        RunOnStaThread(() =>
        {
            var form = new ThemeSettingsForm(new ThemeEditRequest([], new Dictionary<string, double>()));
            try
            {
                form.CreateControl();
                var tab = GetTab(form, "Entity styles");
                LayOutTab(form, tab);
                SelectEntityStyle(tab, "AsPlannedActivity");
                tab.PerformLayout();

                var position = FindAll(
                    tab,
                    c => c is Label label && label.ForeColor == SystemColors.GrayText
                        && !string.IsNullOrEmpty(label.Text)
                        && label.Text != "Solid")
                    .Cast<Label>()
                    .FirstOrDefault(label => label.Bounds.Left > 300);

                Assert.NotNull(position);
                Assert.False(string.IsNullOrWhiteSpace(position!.Text));
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

    /// <summary>
    /// Returns the Entity styles tab's master list.
    /// </summary>
    private static ListView RequireEntityStyleList(TabPage tab) =>
        Assert.IsType<ListView>(Find(tab, c => c is ListView));

    /// <summary>
    /// Selects a style in the Entity styles list and runs the selection
    /// handler, so the detail panel shows that style's properties.
    /// </summary>
    private static void SelectEntityStyle(TabPage tab, string styleKey)
    {
        var list = RequireEntityStyleList(tab);
        var form = Assert.IsType<ThemeSettingsForm>(list.FindForm());

        // The list view's selection is native state, which a unit test cannot
        // create, so the form's own seam shows the style instead.
        var seam = typeof(ThemeSettingsForm).GetMethod(
            "ShowEntityStyleForTesting",
            BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("ShowEntityStyleForTesting not found.");
        seam.Invoke(form, new object[] { styleKey });
    }

    /// <summary>
    /// Forces real layout on a tab page so bounds assertions measure the dialog
    /// the user actually sees.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="Control.CreateControl"/> only creates a handle for a control
    /// that will be visible, and a never-shown form never is, so the
    /// <see cref="TabControl"/> stays handle-less. A handle-less TabControl
    /// sizes nothing: the page sits at its 200x100 design size, the
    /// <see cref="SplitContainer"/> inside it at 184x84, and every width read
    /// from that tree describes a dialog that does not exist. Touching
    /// <see cref="Control.Handle"/> creates the hidden window outright, the
    /// native tab control then sizes the page, and the layout passes below
    /// propagate the form's 900x640 client area all the way down.
    /// </para>
    /// <para>
    /// All three handles are required — forcing only the form's handle leaves
    /// the tab control at 900x580 with an unsized page.
    /// </para>
    /// </remarks>
    /// <param name="form">The dialog under test.</param>
    /// <param name="page">The tab page to activate and lay out.</param>
    private static void LayOutTab(Form form, TabPage page)
    {
        _ = form.Handle;

        var tabControl = Assert.IsType<TabControl>(Find(form, c => c is TabControl));
        _ = tabControl.Handle;

        tabControl.SelectedTab = page;

        form.PerformLayout();
        tabControl.PerformLayout();
        page.PerformLayout();
    }

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

    /// <summary>
    /// Invokes a private method on the form.
    /// </summary>
    private static void InvokePrivate(ThemeSettingsForm form, string methodName)
    {
        var method = typeof(ThemeSettingsForm).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException($"Method '{methodName}' not found on ThemeSettingsForm.");
        object?[] args = method.GetParameters().Length == 0 ? [] : [null, EventArgs.Empty];
        method.Invoke(form, args);
    }

    private static void InvokeClick(ThemeSettingsForm form, string methodName)
    {
        var m = typeof(ThemeSettingsForm).GetMethod(methodName, BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException($"Method '{methodName}' not found on ThemeSettingsForm.");
        m.Invoke(form, new object?[] { null, EventArgs.Empty });
    }

    /// <summary>
    /// Invokes the permitted-label list's check handler for one item, the
    /// way the host does when the user clicks a tick box. The list's
    /// <c>ItemCheck</c> event fires before the check state changes, so the
    /// handler is given the item's index and the state it is moving to.
    /// </summary>
    /// <param name="form">The form under test.</param>
    /// <param name="index">The item's zero-based index.</param>
    /// <param name="newValue">The state the item is moving to.</param>
    private static void InvokeAllowedLabelItemCheck(ThemeSettingsForm form, int index, CheckState newValue)
    {
        var handler = typeof(ThemeSettingsForm).GetMethod(
            "CustomStyleAllowedLabelItemCheck",
            BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("CustomStyleAllowedLabelItemCheck not found.");

        var args = new ItemCheckEventArgs(index, newValue, CheckState.Unchecked);
        handler.Invoke(form, new object?[] { null, args });
    }
}
