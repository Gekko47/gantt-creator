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
}
