
using ExcelDna.Integration.CustomUI;
using GanttCreator.Core;
using GanttCreator.Core.Scene;
using GanttCreator.Office;
using Moq;
using static GanttCreator.AddIn.Tests.RibbonStateTestSupport;

namespace GanttCreator.AddIn.Tests;

/// <summary>
/// The snapshot half of <see cref="RibbonStateService"/>: what one capture reports — the mode, the
/// warnings and validation issues, the export and structural gates, the adopted tab, and the
/// entity-primary defaults — and what the service keeps when a source fails or the workbook changes.
/// </summary>
/// <remarks>
/// Shares the <c>diagnostics-service</c> collection: these facts arm the same session singleton the
/// host-lifecycle facts arm and reset, so xUnit runs the collection's classes sequentially. The
/// constructor and <see cref="Dispose"/> reset it around every fact, which is why the reset stays
/// visible here rather than hiding in a shared base.
/// </remarks>
[Collection("diagnostics-service")]
public sealed class RibbonStateSnapshotTests : IDisposable
{
    public RibbonStateSnapshotTests() => RibbonStateService.Reset();

    public void Dispose()
    {
        RibbonStateService.Reset();
        GC.SuppressFinalize(this);
    }
    [Fact]
    public void Get_catalogue_writer_returns_the_injected_writer()
    {
        // The Initialise-sheet command injects this instance into the
        // workbook initialiser; if the getter lost the writer, the
        // CatalogueWriteRejected evidence record would silently never fire
        // (live observation 2026-10-09T17:04Z).
        var service = new RibbonStateService();
        var writer = new Mock<IConfigCatalogueWriter>().Object;

        service.SetCatalogueWriter(writer);

        Assert.Same(writer, service.GetCatalogueWriter());
    }

    [Fact]
    public void Get_catalogue_writer_is_null_before_injection()
    {
        // The unarmed fallback in RunForExcel relies on null, not on a
        // stale or default instance.
        var service = new RibbonStateService();

        Assert.Null(service.GetCatalogueWriter());
    }

    [Fact]
    public void Current_validation_findings_do_not_replace_last_rendered_scene_warnings()
    {
        var service = new RibbonStateService();
        SceneWarning sceneWarning = new(SceneOwnerId.Chart, "PlotRange", "Chart warning.");
        GanttValidationIssue validationIssue = new(
            3, "Start", "InvalidDate", GanttValidationSeverity.Error, "Enter a valid start date.");

        service.UpdateWarnings([sceneWarning]);
        service.UpdateCurrentValidationIssues([validationIssue]);

        Assert.Equal([sceneWarning], service.GetLastSuccessfulSceneWarnings());
        Assert.Equal([validationIssue], service.GetCurrentValidationIssues());
    }

    [Fact]
    public void A_settings_write_with_no_injected_writer_records_that_it_was_skipped()
    {
        // A half-armed session must be diagnosable: the in-memory state is
        // accepted, but nothing is persisted and nothing is scheduled.
        using var log = new CapturingLog();
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log);
        var service = new RibbonStateService();
        service.SetCommandBoundary(boundary);

        service.SetTimeScale(GanttTimeScale.Quarter);

        Assert.Contains(
            log.Records,
            record => record.Contains("gantt.settings persist skipped=no-writer", StringComparison.Ordinal));
    }

    [Fact]
    public void A_throwing_snapshot_source_keeps_the_previous_state_and_records_the_failure()
    {
        // Sticky snapshots are the policy; silence is not. The record is the
        // only evidence a future snapshot defect leaves behind.
        using var log = new CapturingLog();
        var boundary = new CommandBoundary(presenter: _ => { });
        boundary.SetLog(log);
        var service = new RibbonStateService();
        service.SetCommandBoundary(boundary);
        service.SetLogAvailabilitySource(() => throw new InvalidOperationException("probe"));
        bool openLogBefore = service.GetEnabled(RibbonControlIds.OpenLog);

        service.Refresh();

        Assert.Contains(
            log.Records,
            record => record.Contains("gantt.ribbon snapshot kept-previous", StringComparison.Ordinal));
        Assert.Equal(openLogBefore, service.GetEnabled(RibbonControlIds.OpenLog));
    }

    [Fact]
    public void A_workbook_switch_clears_the_last_rendered_warnings_and_current_issues()
    {
        // R5.8/R5.11: both snapshots are per-workbook. Carrying them across a
        // switch would show another workbook's warnings and stale badges, which
        // the user can never dismiss.
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = false };
        var service = Arm(adapter);
        service.UpdateWarnings([new SceneWarning(SceneOwnerId.Chart, "PlotRange", "Chart warning.")]);
        service.UpdateCurrentValidationIssues(
            [new GanttValidationIssue(3, "Start", "InvalidDate", GanttValidationSeverity.Error, "Enter a valid start date.")]);
        Assert.NotEmpty(service.GetLastSuccessfulSceneWarnings());
        Assert.NotEmpty(service.GetCurrentValidationIssues());
        Assert.NotNull(service.GetLastSuccessfulRenderAtUtc());

        adapter.RaiseWorkbookStateChanged();

        Assert.Empty(service.GetLastSuccessfulSceneWarnings());
        Assert.Empty(service.GetCurrentValidationIssues());
        Assert.Null(service.GetLastSuccessfulRenderAtUtc());
    }

    [Fact]
    public void A_fresh_service_points_every_add_primary_at_its_family_default()
    {
        // R5.18g: before the user has touched a menu, each split button's primary adds the
        // family default its label shows.
        RibbonStateService service = new();

        Assert.Equal(
            GanttEntityType.AsPlannedActivity,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Activity));
        Assert.Equal(
            GanttEntityType.AsPlannedMilestone,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Milestone));
        Assert.Equal(
            GanttEntityType.Delineator,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Element));
    }

    [Fact]
    public void A_menu_pick_moves_only_its_own_family_primary()
    {
        RibbonStateService service = new();

        service.RememberAddedType(GanttEntityType.AsBuiltActivity);

        Assert.Equal(
            GanttEntityType.AsBuiltActivity,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Activity));
        Assert.Equal(
            GanttEntityType.AsPlannedMilestone,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Milestone));
        Assert.Equal(
            GanttEntityType.Delineator,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Element));
    }

    [Fact]
    public void A_remembered_pick_that_is_no_longer_permitted_falls_back_to_the_family_default()
    {
        // The Critical Interval item disables itself when no row can own one. A primary must
        // never stay pointed at an action its own menu greys out, so the projection falls
        // back to the default; an unarmed service reports the item unavailable, which is
        // precisely the condition being pinned.
        RibbonStateService service = new();
        Assert.False(service.GetEnabled(RibbonControlIds.CriticalIntervalMenuItem));

        service.RememberAddedType(GanttEntityType.CriticalInterval);

        Assert.Equal(
            GanttEntityType.AsPlannedActivity,
            service.PrimaryTypeFor(RibbonStateService.AddEntityFamily.Activity));
    }

    [Fact]
    public void The_primary_label_names_what_that_primary_would_add()
    {
        RibbonStateService service = new();

        Assert.Equal(
            "As-Planned Activity",
            RibbonStateService.AddEntityPrimaryLabel(RibbonControlIds.AddActivity, service));
        Assert.Equal(
            "As-Planned Milestone",
            RibbonStateService.AddEntityPrimaryLabel(RibbonControlIds.AddMilestone, service));
        Assert.Equal(
            "Delineator",
            RibbonStateService.AddEntityPrimaryLabel(RibbonControlIds.AddDelineator, service));

        service.RememberAddedType(GanttEntityType.Splitter);

        Assert.Equal(
            "Splitter",
            RibbonStateService.AddEntityPrimaryLabel(RibbonControlIds.AddDelineator, service));
    }

    [Fact]
    public void A_primary_label_for_an_unknown_control_is_empty_rather_than_a_guess()
    {
        RibbonStateService service = new();

        Assert.Equal(string.Empty, RibbonStateService.AddEntityPrimaryLabel(null, service));
        Assert.Equal(
            string.Empty,
            RibbonStateService.AddEntityPrimaryLabel(RibbonControlIds.Remove, service));
    }
    [Fact]
    public void Refresh_reports_first_run_when_no_workbook_is_active()
    {
        // R5.16 positive control for the old `refusal != TableMissing` test:
        // NoActiveWorkbook used to count as initialised, which would make this
        // snapshot Ready and hide the only button a first-run user has.
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = false },
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook));

        Assert.True(service.GetVisible(RibbonControlIds.GetStartedGroup));
        Assert.True(service.GetVisible(RibbonControlIds.DiagnosticsGroup));
        Assert.False(service.GetVisible(RibbonControlIds.DataGroup));
        Assert.False(service.GetVisible(RibbonControlIds.OutlineGroup));
    }

    [Fact]
    public void Refresh_reports_needs_repair_when_configuration_exists_without_a_table()
    {
        // The recovery case: a workbook that carries _GanttCreatorConfig but has
        // lost its table must not be treated as a first run, or Initialise's
        // "use Repair configuration" refusal points at a hidden command.
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Refused(
                GanttTableReadRefusalReason.TableMissing,
                configurationPresent: true));

        Assert.False(service.GetVisible(RibbonControlIds.GetStartedGroup));
        Assert.True(service.GetVisible(RibbonControlIds.DataGroup));
        Assert.True(service.GetEnabled(RibbonControlIds.RepairConfig));
        Assert.False(service.GetEnabled(RibbonControlIds.AddActivity));
    }

    [Fact]
    public void Refresh_reports_ready_when_the_table_is_readable()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([]));

        Assert.False(service.GetVisible(RibbonControlIds.GetStartedGroup));
        Assert.True(service.GetVisible(RibbonControlIds.DataGroup));
        Assert.True(service.GetEnabled(RibbonControlIds.AddActivity));
    }

    [Fact]
    public void Copy_editable_is_enabled_and_visible_on_an_initialised_gantt_sheet()
    {
        // R7.6 D6: the Export commands are gated on the corrected Ready rule — a readable
        // table AND a present configuration — plus the active-Gantt-sheet fact. The
        // configuration is explicit here because Ready requires it (see the NeedsRepair pin).
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([], configurationPresent: true));

        Assert.True(service.GetEnabled(RibbonControlIds.CopyEditable));
        Assert.True(service.GetVisible(RibbonControlIds.ExportGroup));
    }

    [Fact]
    public void Copy_editable_is_disabled_and_hidden_on_a_first_run_workbook()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = false },
            GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.NoActiveWorkbook));

        Assert.False(service.GetEnabled(RibbonControlIds.CopyEditable));
        Assert.False(service.GetVisible(RibbonControlIds.ExportGroup));
    }

    /// <summary>
    /// R7.6 D6: both Export commands are enabled together on a Ready Gantt sheet, and the
    /// group is visible.
    /// </summary>
    [Fact]
    public void Send_to_powerpoint_is_enabled_on_a_ready_gantt_sheet()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([], configurationPresent: true));

        Assert.True(service.GetEnabled(RibbonControlIds.CopyEditable));
        Assert.True(service.GetEnabled(RibbonControlIds.SendToPowerPoint));
        Assert.True(service.GetVisible(RibbonControlIds.ExportGroup));
    }

    /// <summary>
    /// R7.6 D6: the defect the corrected gate removes. A readable table with a missing
    /// configuration is NeedsRepair, where both export commands would refuse at the
    /// configuration read, so the group is hidden and neither button is enabled.
    /// </summary>
    [Fact]
    public void The_export_group_is_hidden_and_both_buttons_disabled_in_needs_repair()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([], configurationPresent: false));

        Assert.False(service.GetVisible(RibbonControlIds.ExportGroup));
        Assert.False(service.GetEnabled(RibbonControlIds.CopyEditable));
        Assert.False(service.GetEnabled(RibbonControlIds.SendToPowerPoint));
    }

    /// <summary>
    /// R7.6 D6: a Ready workbook with another visible sheet active keeps the group visible
    /// but disables both commands, which would paste a chart the user is not looking at.
    /// </summary>
    [Fact]
    public void The_export_buttons_are_disabled_on_another_sheet()
    {
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok(
                [],
                configurationPresent: true,
                activeSheetIsGanttSheet: false));

        Assert.True(service.GetVisible(RibbonControlIds.ExportGroup));
        Assert.False(service.GetEnabled(RibbonControlIds.CopyEditable));
        Assert.False(service.GetEnabled(RibbonControlIds.SendToPowerPoint));
    }

    [Fact]
    public void Structural_buttons_follow_the_row_under_the_active_cell()
    {
        // R5.9b: a child at the bottom of its run offers Promote and Remove,
        // not Demote, Up or Down. The flags come from RowActionAvailability.
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok(
            [
                Row(1, "As-Planned Activity"),
                Row(2, "As-Planned Activity", parentId: NewRowId('a')),
            ]));

        var identifier = new Mock<ISelectionIdentifier>();
        _ = identifier.Setup(i => i.Identify()).Returns(new SelectionFacts(true, true, 2, false, 1, 2));
        service.SetSelectionIdentifier(identifier.Object);
        service.Refresh();

        Assert.True(service.GetEnabled(RibbonControlIds.PromoteToParent));
        Assert.False(service.GetEnabled(RibbonControlIds.DemoteToChild));
        Assert.False(service.GetEnabled(RibbonControlIds.MoveUp));
        Assert.False(service.GetEnabled(RibbonControlIds.MoveDown));
        Assert.True(service.GetEnabled(RibbonControlIds.Remove));

        // A header selection resolves no body row, so Remove greys too.
        _ = identifier.Setup(i => i.Identify()).Returns(new SelectionFacts(true, false, null, true, 0));
        service.Refresh();

        Assert.False(service.GetEnabled(RibbonControlIds.Remove));
        Assert.False(service.GetEnabled(RibbonControlIds.PromoteToParent));
    }

    [Fact]
    public void Chart_commands_are_disabled_when_another_sheet_is_active()
    {
        // R5.16b (ADR-0058): the reader reports that the active sheet is not the
        // sheet that carries tblGanttData, so every chart command greys while the
        // recovery surface stays available from any sheet.
        var service = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([], activeSheetIsGanttSheet: false));

        Assert.False(service.GetEnabled(RibbonControlIds.AddActivity));
        Assert.False(service.GetEnabled(RibbonControlIds.RefreshSheet));
        Assert.False(service.GetEnabled(RibbonControlIds.Warnings));
        Assert.False(service.GetEnabled(RibbonControlIds.ToggleChartLock));
        Assert.False(service.GetEnabled(RibbonControlIds.ToggleCellLock));
        Assert.False(service.GetEnabled(RibbonControlIds.ThemeSettings));
        Assert.False(service.GetEnabled(RibbonControlIds.ChartTitle));
        Assert.False(service.GetEnabled(RibbonControlIds.PlotTimeScale));
        Assert.False(service.GetEnabled(RibbonControlIds.SizeDropdown));
        Assert.False(service.GetEnabled(RibbonControlIds.Margin));

        Assert.True(service.GetEnabled(RibbonControlIds.InitialiseSheet));
        Assert.True(service.GetEnabled(RibbonControlIds.RepairConfig));
        Assert.True(service.GetEnabled(RibbonControlIds.ValidateSheet));
        Assert.True(service.GetEnabled(RibbonControlIds.Diagnostics));
    }

    [Fact]
    public void Expand_and_collapse_are_disabled_on_a_flat_table_and_enabled_with_a_group()
    {
        var flat = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok([Row(1, "As-Planned Activity")]));

        Assert.False(flat.GetEnabled(RibbonControlIds.ExpandAllGroups));
        Assert.False(flat.GetEnabled(RibbonControlIds.CollapseAllGroups));

        var grouped = ArmedWithTableRead(
            new TestExcelApplicationAdapter { ActiveWorkbook = true },
            GanttTableReadOutcome.Ok(
            [
                Row(1, "As-Planned Activity"),
                Row(2, "As-Planned Activity", parentId: NewRowId('a')),
            ]));

        Assert.True(grouped.GetEnabled(RibbonControlIds.ExpandAllGroups));
        Assert.True(grouped.GetEnabled(RibbonControlIds.CollapseAllGroups));
    }

    [Fact]
    public void Refresh_activates_the_add_in_tab_when_the_gantt_sheet_is_active()
    {
        // Owner ruling D-4: every refresh restores the tab, so a chart refresh
        // cannot leave the user looking at the Table Design tab.
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = true }, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);

        // The configuration must be present for the surface to be Ready; a
        // successful read without it is NeedsRepair (R5.17b).
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        service.SetTableReader(reader.Object);

        service.Refresh();

        ribbon.Verify(r => r.ActivateTab(RibbonControlIds.Tab), Times.Once);
    }

    /// <summary>
    /// R5.20: pressing the automatic-refresh toggle restores the add-in tab, for the same reason a
    /// refresh does. Reported from the host on 2026-11-16 — the settings write moves Excel's
    /// selection, which activates the contextual Table Design tab — so the toggle runs the same
    /// ADR-0058 restore rather than leaving the user looking at another tab.
    /// </summary>
    [Fact]
    public void The_auto_refresh_toggle_restores_the_add_in_tab()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = true }, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        service.SetTableReader(reader.Object);
        service.Refresh();

        // The refresh that brings the surface to Ready restores the tab itself, so the calls are
        // cleared: what is asserted is the toggle's own restore.
        ribbon.Invocations.Clear();

        service.SetAutoRefresh(false);

        ribbon.Verify(r => r.ActivateTab(RibbonControlIds.Tab), Times.Once);
    }

    [Fact]
    public void Refresh_does_not_activate_the_tab_when_another_sheet_is_active()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = true }, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok(
                [],
                configurationPresent: true,
                activeSheetIsGanttSheet: false));
        service.SetTableReader(reader.Object);

        service.Refresh();

        ribbon.Verify(r => r.ActivateTab(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public void Refresh_does_not_activate_the_tab_before_the_surface_is_ready()
    {
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(new TestExcelApplicationAdapter { ActiveWorkbook = true }, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Refused(GanttTableReadRefusalReason.TableMissing));
        service.SetTableReader(reader.Object);

        service.Refresh();

        ribbon.Verify(r => r.ActivateTab(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public void A_sheet_activation_refreshes_the_gate_and_restores_the_tab()
    {
        var adapter = new TestExcelApplicationAdapter { ActiveWorkbook = true };
        var ribbon = new Mock<IRibbonUI>();
        var service = Arm(adapter, ribbon: ribbon.Object);
        var reader = new Mock<IGanttTableReader>(MockBehavior.Strict);
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok(
                [],
                configurationPresent: true,
                activeSheetIsGanttSheet: false));
        service.SetTableReader(reader.Object);
        service.Refresh();

        Assert.Equal(1, adapter.SheetActivationSubscribeCount);
        Assert.False(service.GetEnabled(RibbonControlIds.AddActivity));

        // Entering the Gantt sheet: the next read reports it as the active sheet.
        _ = reader.Setup(r => r.Read())
            .Returns(GanttTableReadOutcome.Ok([], configurationPresent: true));
        adapter.RaiseSheetActivated();

        Assert.True(service.GetEnabled(RibbonControlIds.AddActivity));
        ribbon.Verify(r => r.ActivateTab(RibbonControlIds.Tab), Times.Once);
    }
}
