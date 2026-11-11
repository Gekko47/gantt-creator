using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>Runs the explicit configuration integrity check and repair workflow.</summary>
internal static class RepairConfigCommand
{
    /// <summary>The message shown when a healthy workbook has nothing to repair.</summary>
    internal const string NothingToRepairMessage =
        "Nothing to repair: the Gantt table and configuration are already healthy.";

    /// <summary>The message shown when the user declines the recovery confirmation.</summary>
    internal const string CancelledMessage = "Repair cancelled. Nothing was changed.";

    /// <summary>Runs the production repair command in the current Excel session.</summary>
    internal static void RunForExcel()
    {
        object application = ExcelDna.Integration.ExcelDnaUtil.Application;
        Run(
            new ExcelConfigIntegrityChecker(application),
            new ExcelConfigRepairer(application),
            CommandErrorDialog.Confirm,
            CommandErrorDialog.Show,
            new ExcelWorkbookRecovery(
                application,
                trace: CommandBoundary.Instance.WriteTechnicalRecord),
            RibbonStateService.Instance.ScheduleAutomaticRefresh
        );
    }

    /// <summary>Runs one injected inspect → recover → check → plan → confirmation → repair flow.</summary>
    /// <param name="checker">The read-only integrity checker.</param>
    /// <param name="repairer">The plan-constrained repairer.</param>
    /// <param name="confirm">Asks the user to approve a confirm-required change.</param>
    /// <param name="presenter">Receives the summary or typed refusal message.</param>
    /// <param name="recovery">The structural table recovery port, or null to keep the legacy flow.</param>
    /// <param name="onChanged">Schedules one coalesced automatic refresh after a mutation.</param>
    internal static void Run(
        IConfigIntegrityChecker checker,
        IConfigRepairer repairer,
        Func<string, bool> confirm,
        Action<string> presenter,
        IWorkbookRecovery? recovery = null,
        Action? onChanged = null
    )
    {
        ArgumentNullException.ThrowIfNull(checker);
        ArgumentNullException.ThrowIfNull(repairer);
        ArgumentNullException.ThrowIfNull(confirm);
        ArgumentNullException.ThrowIfNull(presenter);

        WorkbookRecoveryInspection? inspection = null;
        WorkbookRecoveryOutcome? recoveryOutcome = null;
        if (recovery is not null)
        {
            inspection = recovery.Inspect();
            if (inspection.Refusal is { } recoveryRefusal)
            {
                Present(presenter, TranslateRecoveryRefusal(recoveryRefusal));
                return;
            }

            if (inspection.NeedsRecovery)
            {
                bool confirmed =
                    !inspection.RequiresConfirmation
                    || confirm(
                        inspection.ConfirmationMessage
                        ?? "Fix will rebuild the Gantt table. Continue?");
                if (!confirmed)
                {
                    Present(presenter, CancelledMessage);
                    return;
                }

                recoveryOutcome = recovery.Recover(inspection, confirmed);
                if (recoveryOutcome.Refusal is { } failedRecovery)
                {
                    Present(presenter, TranslateRecoveryRefusal(failedRecovery));
                    return;
                }
            }
        }

        ConfigIntegrityCheckOutcome check = checker.Check();
        if (check.Refusal is not null)
        {
            Present(presenter, TranslateRefusal(check.Refusal.Value));
            return;
        }

        ConfigIntegrityPlan plan = ConfigIntegrityPlan.Build(check.Findings);
        if (plan.Findings.Count == 0)
        {
            Present(presenter, BuildNoRepairSummary(inspection, recoveryOutcome));
            ScheduleRefreshIfChanged(onChanged, recoveryOutcome, repaired: 0);
            return;
        }

        int confirmCount = plan.ConfirmRequiredFindings.Count;
        if (plan.RefusedFindings.Count > 0)
        {
            Present(
                presenter,
                BuildSummaryWithRecovery(
                    inspection,
                    recoveryOutcome,
                    0,
                    0,
                    plan.RefusedFindings.Count,
                    plan.QuarantinedFindings.Count));
            ScheduleRefreshIfChanged(onChanged, recoveryOutcome, repaired: 0);
            return;
        }

        bool confirmedRepair =
            confirmCount == 0
            || confirm(BuildConfirmation(confirmCount, plan.QuarantinedFindings.Count));
        ConfigRepairOutcome outcome = repairer.Repair(plan, confirmedRepair);
        if (outcome.Refusal is not null)
        {
            Present(presenter, TranslateRefusal(outcome.Refusal.Value));
            ScheduleRefreshIfChanged(onChanged, recoveryOutcome, repaired: 0);
            return;
        }

        Present(
            presenter,
            BuildSummaryWithRecovery(
                inspection,
                recoveryOutcome,
                outcome.RepairedCount,
                outcome.ConfirmationSkippedCount,
                outcome.RefusedCount,
                outcome.QuarantinedCount));
        ScheduleRefreshIfChanged(onChanged, recoveryOutcome, outcome.RepairedCount);
    }

    private static void ScheduleRefreshIfChanged(
        Action? onChanged,
        WorkbookRecoveryOutcome? recoveryOutcome,
        int repaired)
    {
        if (onChanged is null)
        {
            return;
        }

        if (recoveryOutcome?.Changed == true || repaired > 0)
        {
            onChanged();
        }
    }

    private static string BuildNoRepairSummary(
        WorkbookRecoveryInspection? inspection,
        WorkbookRecoveryOutcome? recoveryOutcome) =>
        recoveryOutcome?.Changed == true
            ? BuildRecoveryPrefix(inspection, recoveryOutcome)
                + "The configuration needed no repair."
            : NothingToRepairMessage;

    private static string BuildSummaryWithRecovery(
        WorkbookRecoveryInspection? inspection,
        WorkbookRecoveryOutcome? recoveryOutcome,
        int repaired,
        int skipped,
        int refused,
        int quarantined)
    {
        string prefix = BuildRecoveryPrefix(inspection, recoveryOutcome);

        return repaired == 0 && skipped == 0 && refused == 0 && quarantined == 0
            ? prefix + "The configuration needed no repair."
            : prefix + BuildSummary(repaired, skipped, refused, quarantined);
    }

    private static string BuildRecoveryPrefix(
        WorkbookRecoveryInspection? inspection,
        WorkbookRecoveryOutcome? recoveryOutcome)
    {
        List<string> parts = [];
        if (recoveryOutcome?.GanttSheetCreated == true)
        {
            parts.Add("Created a new empty Gantt worksheet from the stored configuration.");
        }

        if (recoveryOutcome?.ConfigurationCreated == true)
        {
            parts.Add("Recreated _GanttCreatorConfig from code defaults.");
        }

        if (recoveryOutcome?.ConfigurationRenamed == true)
        {
            parts.Add("Restored the _GanttCreatorConfig worksheet name.");
        }

        if (recoveryOutcome?.TableRenamed == true)
        {
            parts.Add("Renamed the Gantt table to tblGanttData.");
        }

        if (recoveryOutcome?.TableRebuilt == true)
        {
            string sheet = inspection?.TargetSheetName ?? "the Gantt worksheet";
            parts.Add(
                string.Create(
                    System.Globalization.CultureInfo.InvariantCulture,
                    $"Rebuilt tblGanttData on '{sheet}' and adopted {recoveryOutcome.AdoptedRowCount} row(s)."));
        }

        if (recoveryOutcome?.TableStructureRepaired == true)
        {
            parts.Add("Restored the Gantt table structure.");
        }

        return parts.Count == 0 ? string.Empty : string.Join(" ", parts) + " ";
    }

    private static string BuildConfirmation(int confirmCount, int quarantineCount)
    {
        string quarantine =
            quarantineCount == 0
                ? string.Empty
                : $" {quarantineCount} user style row(s) are quarantined and will not be discarded silently.";
        return $"Repair configuration will change {confirmCount} add-in-owned or user-data item(s).{quarantine} Continue?";
    }

    private static string BuildSummary(int repaired, int skipped, int refused, int quarantined) =>
        $"Configuration repair complete: {repaired} repaired, {skipped} skipped, {refused} refused, {quarantined} quarantined.";

    private static string TranslateRecoveryRefusal(WorkbookRecoveryRefusalReason refusal) =>
        refusal switch
        {
            WorkbookRecoveryRefusalReason.NoActiveWorkbook =>
                "Gantt Creator needs an open workbook before it can repair the Gantt table.",
            WorkbookRecoveryRefusalReason.TargetProtected =>
                "The Gantt worksheet or workbook structure is protected, so the table was not repaired. Unprotect it and try Fix again.",
            WorkbookRecoveryRefusalReason.TargetMissing =>
                "Fix could not find the Gantt worksheet. It looks for exactly one worksheet named 'Gantt Data'.",
            WorkbookRecoveryRefusalReason.TargetAmbiguous =>
                "Fix found more than one worksheet named 'Gantt Data', so it could not choose one. Rename the extra sheet and try Fix again.",
            WorkbookRecoveryRefusalReason.TargetHasForeignTable =>
                "The Gantt worksheet already carries a different table, so Fix did not rebuild it. Rename that table to tblGanttData or remove it, then try Fix again.",
            WorkbookRecoveryRefusalReason.HeaderMisaligned =>
                "The Gantt worksheet's header row no longer matches the Gantt Creator schema, so Fix did not overwrite it. Restore the header row or reinitialise a copy of the workbook.",
            WorkbookRecoveryRefusalReason.TableStructureRefused =>
                "The Gantt table's columns are damaged in a way Fix cannot repair without guessing, so nothing was changed. Check the table's header row.",
            WorkbookRecoveryRefusalReason.TableCreateFailed =>
                "Excel refused to create tblGanttData, so the table was not rebuilt. Try again; if it keeps failing, see the Diagnostics dialog.",
            WorkbookRecoveryRefusalReason.HeaderWriteFailed =>
                "Excel refused a header write, so the table was not repaired. Try again; if it keeps failing, see the Diagnostics dialog.",
            WorkbookRecoveryRefusalReason.VerificationFailed =>
                "Fix created the table but could not read it back, so it rolled the change back. Check the sheet and try again.",
            WorkbookRecoveryRefusalReason.DateSystemUnsupported =>
                "This workbook uses the 1904 date system, which Gantt Creator does not support.",
            WorkbookRecoveryRefusalReason.ConfirmationDeclined =>
                CancelledMessage,
            WorkbookRecoveryRefusalReason.ConfigurationNameCollision =>
                "A sheet named '_GanttCreatorConfig' already exists but is not Gantt Creator's configuration, so nothing was changed. Rename that sheet and try Fix again.",
            WorkbookRecoveryRefusalReason.ConfigurationAmbiguous =>
                "Fix found more than one sheet that could be the Gantt Creator configuration, so it changed nothing. Remove the extra sheet and try Fix again.",
            WorkbookRecoveryRefusalReason.ConfigurationCandidateAmbiguous =>
                "A visible sheet carries Gantt Creator configuration tables, so Fix did not create a second configuration. Rename or remove that sheet and try Fix again.",
            WorkbookRecoveryRefusalReason.ConfigurationUnrecognised =>
                "A hidden sheet could be a damaged Gantt Creator configuration, so Fix did not create a second one. Restore its name to '_GanttCreatorConfig' in the Visual Basic Editor and try Fix again.",
            WorkbookRecoveryRefusalReason.ConfigurationRenameFailed =>
                "Excel refused to rename the configuration worksheet, so nothing was changed. Rename it to '_GanttCreatorConfig' in the Visual Basic Editor and try Fix again.",
            WorkbookRecoveryRefusalReason.TableRenameFailed =>
                "Excel refused to rename the Gantt table, so nothing was changed. Rename the table to tblGanttData and try Fix again.",
            WorkbookRecoveryRefusalReason.ConfigurationCreateFailed =>
                "Excel refused to create the Gantt Creator configuration worksheet, so nothing was changed. Try again; if it keeps failing, see the Diagnostics dialog.",
            WorkbookRecoveryRefusalReason.ConfigurationWriteFailed =>
                "The configuration worksheet was created but its catalogues could not be written, so the change was rolled back. Try again; if it keeps failing, see the Diagnostics dialog.",
            WorkbookRecoveryRefusalReason.GanttSheetCreateFailed =>
                "Excel refused to create the Gantt worksheet, so nothing was changed. Try again; if it keeps failing, see the Diagnostics dialog.",
            _ => "Gantt Creator could not repair the Gantt table. See the Diagnostics dialog.",
        };

    private static string TranslateRefusal(ConfigIntegrityCheckRefusalReason refusal) =>
        refusal == ConfigIntegrityCheckRefusalReason.NoActiveWorkbook
            ? "Gantt Creator needs an open workbook before it can check configuration."
            : "Gantt Creator could not check configuration.";

    private static string TranslateRefusal(ConfigRepairRefusalReason refusal) =>
        refusal switch
        {
            ConfigRepairRefusalReason.NoActiveWorkbook => "Gantt Creator needs an open workbook before it can repair configuration.",
            ConfigRepairRefusalReason.TargetProtected =>
                "The workbook or target worksheet is protected, so no configuration was changed. Unprotect it and try again.",
            ConfigRepairRefusalReason.PlanRefused => "Configuration repair was refused because unsafe or unsupported damage was found.",
            ConfigRepairRefusalReason.CatalogueWriteRefused =>
                "Gantt Creator could not safely write the configuration catalogues. No further changes were made.",
            ConfigRepairRefusalReason.VisibilityRepairFailed =>
                "Gantt Creator could not repair the configuration worksheet visibility. No further changes were made.",
            ConfigRepairRefusalReason.PlotAnchorRepairFailed =>
                "Gantt Creator could not repair the plot anchor. No further changes were made.",
            ConfigRepairRefusalReason.IdentityRepairFailed =>
                "Gantt Creator could not safely repair row identity. No further changes were made.",
            ConfigRepairRefusalReason.TypeOptionsUnavailable =>
                "Configuration was repaired, but the Type dropdown could not be prepared. No further changes were made.",
            _ => "Gantt Creator could not repair configuration. See the Diagnostics dialog.",
        };

    private static void Present(Action<string> presenter, string message)
    {
#pragma warning disable CA1031
        try
        {
            presenter(message);
        }
        catch
        {
            // A dialog failure must not escape into Excel.
        }
#pragma warning restore CA1031
    }
}
