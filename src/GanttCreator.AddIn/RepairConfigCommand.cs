using GanttCreator.Core.ConfigIntegrity;
using GanttCreator.Office;

namespace GanttCreator.AddIn;

/// <summary>Runs the explicit configuration integrity check and repair workflow.</summary>
internal static class RepairConfigCommand
{
    /// <summary>Runs the production repair command in the current Excel session.</summary>
    internal static void RunForExcel() =>
        Run(
            new ExcelConfigIntegrityChecker(ExcelDna.Integration.ExcelDnaUtil.Application),
            new ExcelConfigRepairer(ExcelDna.Integration.ExcelDnaUtil.Application),
            CommandErrorDialog.Confirm,
            CommandErrorDialog.Show
        );

    /// <summary>Runs one injected check → plan → confirmation → repair flow.</summary>
    /// <param name="checker">The read-only integrity checker.</param>
    /// <param name="repairer">The plan-constrained repairer.</param>
    /// <param name="confirm">Asks the user to approve a confirm-required change.</param>
    /// <param name="presenter">Receives the count-only summary or typed refusal message.</param>
    internal static void Run(
        IConfigIntegrityChecker checker,
        IConfigRepairer repairer,
        Func<string, bool> confirm,
        Action<string> presenter
    )
    {
        ArgumentNullException.ThrowIfNull(checker);
        ArgumentNullException.ThrowIfNull(repairer);
        ArgumentNullException.ThrowIfNull(confirm);
        ArgumentNullException.ThrowIfNull(presenter);

        ConfigIntegrityCheckOutcome check = checker.Check();
        if (check.Refusal is not null)
        {
            Present(presenter, TranslateRefusal(check.Refusal.Value));
            return;
        }

        var plan = ConfigIntegrityPlan.Build(check.Findings);
        if (plan.Findings.Count == 0)
        {
            return;
        }

        var confirmCount = plan.ConfirmRequiredFindings.Count;
        if (plan.RefusedFindings.Count > 0)
        {
            Present(presenter, BuildSummary(0, 0, plan.RefusedFindings.Count, plan.QuarantinedFindings.Count));
            return;
        }

        var confirmed = confirmCount == 0 || confirm(BuildConfirmation(confirmCount, plan.QuarantinedFindings.Count));
        ConfigRepairOutcome outcome = repairer.Repair(plan, confirmed);
        if (outcome.Refusal is not null)
        {
            Present(presenter, TranslateRefusal(outcome.Refusal.Value));
            return;
        }

        Present(
            presenter,
            BuildSummary(outcome.RepairedCount, outcome.ConfirmationSkippedCount, outcome.RefusedCount, outcome.QuarantinedCount)
        );
    }

    private static string BuildConfirmation(int confirmCount, int quarantineCount)
    {
        var quarantine =
            quarantineCount == 0
                ? string.Empty
                : $" {quarantineCount} user style row(s) are quarantined and will not be discarded silently.";
        return $"Repair configuration will change {confirmCount} add-in-owned or user-data item(s).{quarantine} Continue?";
    }

    private static string BuildSummary(int repaired, int skipped, int refused, int quarantined) =>
        $"Configuration repair complete: {repaired} repaired, {skipped} skipped, {refused} refused, {quarantined} quarantined.";

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
