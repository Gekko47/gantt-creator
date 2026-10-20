using System.Globalization;
using GanttCreator.Core;
using GanttCreator.Core.Logging;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live <see cref="IConfigCatalogueWriter"/> over the Excel application
/// object. Materialises the five code-owned catalogue tables (ADR-0007 D2)
/// on the <c>_GanttCreatorConfig</c> worksheet; preserves user-authored
/// style rows and present setting values on regeneration (ADR-0007 D4).
/// </summary>
/// <param name="application">
/// The Excel application object (for example <c>ExcelDnaUtil.Application</c>), or
/// <see langword="null"/> when the host supplied none (unit tests, non-Excel
/// host). A foreign object fails the interface cast and degrades to the
/// no-active-workbook refusal with no mutation.
/// </param>
/// <param name="protectionGuard">The shared read-only workbook-protection guard.</param>
/// <param name="log">
/// Optional rolling log that receives one redacted evidence record when
/// the host rejects the write (see
/// <see cref="ConfigWriteRefusalReason.HostRejected"/>). Absent (or
/// failed) logging degrades to no record, never to a thrown error.
/// </param>
/// <remarks>
/// <para>
/// COM ownership: every proxy is held in a local and used without chained
/// member expressions, matching <see cref="ExcelWorkbookInitialiser"/>. The
/// <c>internal virtual</c> accessors isolate the COM parameterised
/// properties (indexers and <c>ListObjects.Add</c>), so contract tests can
/// substitute them (CS0855); the real behaviour is exercised by the tagged
/// live-Office integration test.
/// </para>
/// <para>
/// Mutation order (ADR-0007 D4): all read-only checks and the preservation
/// capture first, then per-table delete-and-recreate. <c>tblGanttTypes</c>,
/// <c>tblGanttMetrics</c>, and <c>tblGanttConfig</c> hold no user content
/// and regenerate outright; <c>tblGanttStyles</c> preserves non-built-in
/// rows and <c>tblGanttSettings</c> preserves present values, appending
/// missing defaults. The port writes only on the configuration worksheet —
/// never schedule data, never the visible sheet, never a second helper
/// sheet.
/// </para>
/// </remarks>
public class ExcelConfigCatalogueWriter(
    object? application,
    IWorksheetProtectionGuard? protectionGuard = null,
    IRollingLog? log = null) : IConfigCatalogueWriter
{
    private readonly Application? _application = application as Application;
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);

    // Evidence sink for a host-rejected write. The broad catch in Write()
    // converts the host exception into a typed refusal, and the command
    // boundary logs nothing for a typed refusal — without this record the
    // original error (type, HRESULT, stack) would be discarded and the live
    // failure could not be diagnosed (live defect chain GC-20261009).
    private readonly IRollingLog? _log = log;

    /// <summary>
    /// The preservation data captured before any mutation: existing setting
    /// values, user-authored style rows, and the workbook ID.
    /// </summary>
    private sealed record CataloguePreservation(
        IReadOnlyDictionary<string, string> Settings,
        IReadOnlyList<object?[]> UserStyleRows,
        string? WorkbookId);

    /// <summary>
    /// One catalogue table's state as it was before <see cref="Write"/> ran,
    /// captured read-only so a mid-sequence host failure can be undone.
    /// </summary>
    /// <param name="Anchor">The anchor cell address the table is written at.</param>
    /// <param name="TableName">The contract table name.</param>
    /// <param name="Existed">
    /// Whether the table was present before this call. A table that did not
    /// exist is removed on rollback; one that did is rewritten from
    /// <paramref name="Headers"/> and <paramref name="Body"/>.
    /// </param>
    /// <param name="Headers">The prior header row, empty when absent.</param>
    /// <param name="Body">The prior body rows, empty when absent.</param>
    private sealed record TableSnapshot(
        string Anchor,
        string TableName,
        bool Existed,
        string[] Headers,
        List<object?[]> Body);

    /// <summary>
    /// The five contract table names in write order, used to capture prior
    /// state before the first mutation.
    /// </summary>
    private static readonly string[] _contractTableNames =
    [
        GanttCatalogues.TypesTableName,
        GanttCatalogues.StylesTableName,
        GanttCatalogues.MetricsTableName,
        GanttCatalogues.SettingsTableName,
        GanttCatalogues.ConfigTableName,
    ];

    /// <inheritdoc />
    public ConfigWriteOutcome Write()
    {
        Application? application = _application;
        if (application is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        // Resolve the actual mutation target. The writer mutates only the
        // configuration worksheet, so QueryTarget(config) is the authoritative
        // check: it reports workbook-structure protection and the
        // configuration sheet's own contents protection. Query() would report
        // the *active* sheet, whose protection says nothing about whether this
        // write may proceed, so it is not consulted here.
        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ConfigSheetMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(config);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigWriteRefusalReason.NoActiveWorkbook
                    : ConfigWriteRefusalReason.TargetProtected);
        }

        // Read-only phase: capture the user content the write must preserve
        // (ADR-0007 D4) before any table is touched.
        CataloguePreservation? preservation = ReadPreservation(config);
        if (preservation is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.CataloguePreservationInvalid);
        }

        // Read-only capture of every table's prior state, before the first
        // mutation, so a host failure partway through the sequence can be
        // undone. This is the transactional counterpart of the per-step
        // rollback in ExcelWorkbookInitialiser.
        List<TableSnapshot> snapshots = CapturePriorState(config);

        // The extent each table was written over, recorded as it is written so
        // a rollback can clear the cells of a table that did not exist before
        // (WriteOrReplaceTable writes the cell values before it creates the
        // list object, so a throw can leave cells behind with no table).
        Dictionary<string, Excel.Range> written = new(StringComparer.Ordinal);

        // Mutation phase: per-table delete-and-recreate under one
        // DisplayAlerts save/restore (table deletion prompts).
        bool original = _application!.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            WriteContractTables(config, preservation, written);
        }
#pragma warning disable CA1031 // Deliberately broad: see the catch body.
        catch (Exception exception)
#pragma warning restore CA1031
        {
            // A host refusal reaching managed code is reported either as a
            // COMException or, when it crosses the interop boundary, as an
            // InvalidOperationException, and the host chooses which. The catch
            // is broad so the rollback runs either way; it is not an
            // ignore-and-continue, because the rollback is the whole point and
            // a failure inside the rollback propagates rather than being
            // swallowed.
            //
            // The evidence record runs BEFORE the rollback: the original
            // exception is about to be discarded in favour of the typed
            // refusal, and a rollback failure would replace it with its own
            // exception — so this is the only point the real cause can be
            // captured. The log write never throws (RollingLog latches its
            // own failures), matching the DiagnosticsService precedent.
            WriteHostRejectedEvidence(exception, written);
            RollBack(config, snapshots, written);
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected);
        }
        finally
        {
            _application.DisplayAlerts = original;
        }

        return ConfigWriteOutcome.Ok();
    }

    /// <inheritdoc />
    public ConfigWriteOutcome WriteTheme(ThemeEditRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (!ThemeEditRequest.TryValidate(request, out _, out _))
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ThemeValidationFailed);
        }

        Application? application = _application;
        if (application is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ConfigSheetMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(config);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigWriteRefusalReason.NoActiveWorkbook
                    : ConfigWriteRefusalReason.TargetProtected);
        }

        // Read-only phase first: capture the two themed tables' prior state,
        // plus the Config table's prior state so the hash update can be
        // rolled back on host failure. The call mutates nothing until the
        // mutation phase below.
        ListObject? stylesTable = FindTable(config, GanttCatalogues.StylesTableName);
        ListObject? metricsTable = FindTable(config, GanttCatalogues.MetricsTableName);
        if (stylesTable is null || metricsTable is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.CataloguePreservationInvalid);
        }

        List<TableSnapshot> snapshots = CaptureThemedPriorState(config);

        // Capture Config table's prior state for rollback of the hash update.
        ListObject? configTable = FindTable(config, GanttCatalogues.ConfigTableName);
        snapshots.Add(
            configTable is null
                ? new TableSnapshot(
                    GanttCatalogues.ConfigAnchor,
                    GanttCatalogues.ConfigTableName,
                    false,
                    [],
                    [])
                : new TableSnapshot(
                    GanttCatalogues.ConfigAnchor,
                    GanttCatalogues.ConfigTableName,
                    true,
                    [.. ReadHeaders(configTable)],
                    ReadBodyRows(configTable)));

        CataloguePreservation? preservation = ReadPreservation(config);
        if (preservation is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.CataloguePreservationInvalid);
        }

        Dictionary<string, Excel.Range> written = new(StringComparer.Ordinal);
        bool original = _application!.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            _ = WriteOrReplaceTable(
                config,
                GanttCatalogues.StylesAnchor,
                GanttCatalogues.StylesTableName,
                GanttCatalogues.StylesHeaders,
                BuildThemedStyleRows(request, preservation),
                written);
            _ = WriteOrReplaceTable(
                config,
                GanttCatalogues.MetricsAnchor,
                GanttCatalogues.MetricsTableName,
                GanttCatalogues.MetricsHeaders,
                BuildThemedMetricRows(request),
                written);

            // The chart-chrome colours are settings, so they are written to
            // the settings table alongside the themed styles and metrics.
            // The write is a whole-table rewrite, so the unedited settings
            // are preserved from the sheet rather than regenerated: the
            // chart title and the scale the user set must survive a colour
            // change.
            _ = WriteOrReplaceTable(
                config,
                GanttCatalogues.SettingsAnchor,
                GanttCatalogues.SettingsTableName,
                GanttCatalogues.SettingsHeaders,
                BuildThemedSettingRows(request, preservation),
                written);

            // Recompute and persist the catalogue hash so the next Read()
            // validates against the updated Styles and Metrics content.
            _ = WriteOrReplaceTable(
                config,
                GanttCatalogues.ConfigAnchor,
                GanttCatalogues.ConfigTableName,
                GanttCatalogues.ConfigHeaders,
                BuildConfigRows(preservation.WorkbookId),
                written);
        }
#pragma warning disable CA1031 // Deliberately broad: see the Write catch body.
        catch (Exception exception)
#pragma warning restore CA1031
        {
            WriteHostRejectedEvidence(exception, written);
            RollBack(config, snapshots, written);
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected);
        }
        finally
        {
            _application.DisplayAlerts = original;
        }

        return ConfigWriteOutcome.Ok();
    }

    /// <summary>
    /// Writes one redacted evidence record for a host-rejected write: the
    /// exception type, HRESULT, the contract tables whose extents were
    /// recorded before the refusal (in write order — the last name is where
    /// the sequence stopped), and the full exception detail. Without this
    /// record the typed <see cref="ConfigWriteRefusalReason.HostRejected"/>
    /// refusal discards the original failure and the live cause cannot be
    /// diagnosed. A missing or failed log degrades to no record.
    /// </summary>
    /// <param name="exception">The host exception being converted to a refusal.</param>
    /// <param name="written">
    /// The extents recorded before the failure, keyed by contract table name.
    /// </param>
    private void WriteHostRejectedEvidence(
        Exception exception,
        Dictionary<string, Excel.Range> written)
    {
        string progress = string.Join(",", _contractTableNames.Where(written.ContainsKey));
        _log?.Write(
            "CatalogueWriteRejected exception={0} hresult=0x{1:X8} written={2} detail={3}",
            exception.GetType().FullName,
            exception.HResult,
            progress,
            exception);
    }

    /// <inheritdoc />
    public ConfigWriteOutcome WriteSettings(IReadOnlyDictionary<string, string> settings)
    {
        ArgumentNullException.ThrowIfNull(settings);

        Application? application = _application;
        if (application is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Workbook? workbook = application.ActiveWorkbook;
        if (workbook is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.NoActiveWorkbook);
        }

        Sheets sheets = workbook.Sheets;
        Worksheet? config = FindConfigSheet(sheets);
        if (config is null)
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.ConfigSheetMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(config);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return ConfigWriteOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? ConfigWriteRefusalReason.NoActiveWorkbook
                    : ConfigWriteRefusalReason.TargetProtected);
        }

        // Read-only phase: capture present values before any mutation so keys
        // outside the supplied update survive (ADR-0007 D4). An absent table
        // contributes nothing and the supplied values fill the approved keys;
        // a corrupt table cannot be preserved safely.
        if (!TryReadKeyValues(config, GanttCatalogues.SettingsTableName, out Dictionary<string, string> existing))
        {
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.CataloguePreservationInvalid);
        }

        Dictionary<string, string> merged = new(existing, StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> entry in settings)
        {
            merged[entry.Key] = entry.Value;
        }

        // Single-table counterpart of the five-table transactional write:
        // capture the settings table's prior state before the first mutation
        // so a host failure partway through is undone, not half-applied.
        ListObject? priorTable = FindTable(config, GanttCatalogues.SettingsTableName);
        TableSnapshot snapshot = priorTable is null
            ? new TableSnapshot(
                GanttCatalogues.SettingsAnchor,
                GanttCatalogues.SettingsTableName,
                false,
                [],
                [])
            : new TableSnapshot(
                GanttCatalogues.SettingsAnchor,
                GanttCatalogues.SettingsTableName,
                true,
                [.. ReadHeaders(priorTable)],
                ReadBodyRows(priorTable));
        List<TableSnapshot> snapshots = [snapshot];
        Dictionary<string, Excel.Range> written = new(StringComparer.Ordinal);

        bool original = _application!.DisplayAlerts;
        try
        {
            _application.DisplayAlerts = false;
            _ = WriteOrReplaceTable(
                config,
                GanttCatalogues.SettingsAnchor,
                GanttCatalogues.SettingsTableName,
                GanttCatalogues.SettingsHeaders,
                BuildSettingRows(merged),
                written);
        }
#pragma warning disable CA1031
        catch (Exception)
#pragma warning restore CA1031
        {
            // Same host-refusal contract as Write: the rollback is the point,
            // and a failure inside it propagates rather than being swallowed.
            RollBack(config, snapshots, written);
            return ConfigWriteOutcome.Refused(ConfigWriteRefusalReason.HostRejected);
        }
        finally
        {
            _application.DisplayAlerts = original;
        }

        return ConfigWriteOutcome.Ok();
    }

    /// <summary>
    /// Writes the five contract tables in order, stopping at the first
    /// failure, and records each table's extent as it is written.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="preservation">The captured user content to preserve.</param>
    /// <param name="written">Receives each written table's extent range.</param>
    private void WriteContractTables(
        Worksheet config,
        CataloguePreservation preservation,
        Dictionary<string, Excel.Range> written)
    {
        _ = WriteOrReplaceTable(
            config,
            GanttCatalogues.TypesAnchor,
            GanttCatalogues.TypesTableName,
            GanttCatalogues.TypesHeaders,
            BuildTypeRows(),
            written);
        _ = WriteOrReplaceTable(
            config,
            GanttCatalogues.StylesAnchor,
            GanttCatalogues.StylesTableName,
            GanttCatalogues.StylesHeaders,
            BuildStyleRows(preservation.UserStyleRows),
            written);
        _ = WriteOrReplaceTable(
            config,
            GanttCatalogues.MetricsAnchor,
            GanttCatalogues.MetricsTableName,
            GanttCatalogues.MetricsHeaders,
            BuildMetricRows(),
            written);
        _ = WriteOrReplaceTable(
            config,
            GanttCatalogues.SettingsAnchor,
            GanttCatalogues.SettingsTableName,
            GanttCatalogues.SettingsHeaders,
            BuildSettingRows(preservation.Settings),
            written);
        _ = WriteOrReplaceTable(
            config,
            GanttCatalogues.ConfigAnchor,
            GanttCatalogues.ConfigTableName,
            GanttCatalogues.ConfigHeaders,
            BuildConfigRows(preservation.WorkbookId),
            written);
    }

    /// <summary>
    /// Captures every contract table's prior state, read-only, before the
    /// first mutation. A table that is absent is recorded as absent so the
    /// rollback removes it rather than leaving a newly created one behind.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>One snapshot per contract table, in write order.</returns>
    private List<TableSnapshot> CapturePriorState(Worksheet config)
    {
        List<TableSnapshot> snapshots = new List<TableSnapshot>(_contractTableNames.Length);
        foreach (string tableName in _contractTableNames)
        {
            ListObject? table = FindTable(config, tableName);
            snapshots.Add(
                table is null
                    ? new TableSnapshot(AnchorFor(tableName), tableName, false, [], [])
                    : new TableSnapshot(
                        AnchorFor(tableName),
                        tableName,
                        true,
                        [.. ReadHeaders(table)],
                        ReadBodyRows(table)));
        }

        return snapshots;
    }

    /// <summary>
    /// Captures the two themed tables' prior state, read-only, before the
    /// first mutation. Shares <see cref="RollBack"/> with the five-table path.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>One snapshot per themed table, in write order.</returns>
    private List<TableSnapshot> CaptureThemedPriorState(Worksheet config)
    {
        string[] themed =
        [
            GanttCatalogues.StylesTableName,
            GanttCatalogues.MetricsTableName,
            GanttCatalogues.SettingsTableName,
        ];
        List<TableSnapshot> snapshots = new List<TableSnapshot>(themed.Length);
        foreach (string tableName in themed)
        {
            ListObject? table = FindTable(config, tableName);
            snapshots.Add(
                table is null
                    ? new TableSnapshot(AnchorFor(tableName), tableName, false, [], [])
                    : new TableSnapshot(
                        AnchorFor(tableName),
                        tableName,
                        true,
                        [.. ReadHeaders(table)],
                        ReadBodyRows(table)));
        }

        return snapshots;
    }

    /// <summary>
    /// Builds themed style rows: catalogue order, edited colours applied to
    /// matching built-ins, preserved user rows appended verbatim.
    /// </summary>
    /// <param name="request">The validated theme edit.</param>
    /// <param name="preservation">The captured user content to preserve.</param>
    /// <returns>The style payload rows.</returns>
    private static List<object?[]> BuildThemedStyleRows(
        ThemeEditRequest request,
        CataloguePreservation preservation)
    {
        Dictionary<string, ThemeStyleEdit> edits = new(StringComparer.Ordinal);
        foreach (ThemeStyleEdit edit in request.StyleEdits)
        {
            edits[edit.StyleKey] = edit;
        }

        // Parse existing user custom style rows into mutable dictionaries.
        Dictionary<string, List<object?>> customStyles = new Dictionary<string, List<object?>>(StringComparer.Ordinal);
        foreach (object?[] row in preservation.UserStyleRows)
        {
            string key = ToText(row[0]);
            customStyles[key] = [.. row.Cast<object?>()];
        }

        // Apply custom style edits: create, rename, delete.
        foreach (CustomStyleEdit custom in request.CustomStyleEdits ?? [])
        {
            switch (custom.Operation)
            {
                case CustomStyleOperation.Create:
                    if (customStyles.ContainsKey(custom.StyleKey))
                    {
                        continue;
                    }
                    customStyles[custom.StyleKey] = BuildCustomStyleRow(custom);
                    break;

                case CustomStyleOperation.Rename:
                    if (!customStyles.TryGetValue(custom.StyleKey, out List<object?>? existing))
                    {
                        continue;
                    }
                    _ = customStyles.Remove(custom.StyleKey);
                    List<object?> renamed = BuildCustomStyleRow(custom, existing);
                    customStyles[custom.NewStyleKey ?? custom.StyleKey] = renamed;
                    break;

                case CustomStyleOperation.Update:
                    if (!customStyles.TryGetValue(custom.StyleKey, out List<object?>? existingUpdate))
                    {
                        continue;
                    }
                    List<object?> updated = BuildCustomStyleRow(custom, existingUpdate);
                    customStyles[custom.StyleKey] = updated;
                    break;

                case CustomStyleOperation.Delete:
                    _ = customStyles.Remove(custom.StyleKey);
                    break;
                default:
                    break;
            }
        }

        List<object?[]> rows = [];
        foreach (GanttStylePreset preset in GanttCatalogues.StylePresets)
        {
            if (edits.TryGetValue(preset.StyleKey, out ThemeStyleEdit? edit))
            {
                rows.Add(
                [
                    preset.StyleKey,
                    preset.DisplayName,
                    edit.FillColour,
                    edit.StrokeColour,
                    preset.HatchPattern.ToString(),
                    preset.HatchPitchPt,
                    preset.HatchLinePt,
                    edit.TextColour,
                    preset.StandardOutlinePt,
                    preset.ActivityHeightPt,
                    preset.MilestoneSizePt,
                    preset.DefaultLabelPosition.ToString(),
                    string.Join(
                        " ",
                        preset.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)),
                    preset.ColourCapability.ToString(),
                    (GanttStyleDashPolicy.IsUserEditable(preset.StyleKey)
                        ? edit.StrokeDash ?? preset.StrokeDash
                        : preset.StrokeDash).ToString(),
                ]);
            }
            else
            {
                rows.Add(
                [
                    preset.StyleKey,
                    preset.DisplayName,
                    preset.FillColour,
                    preset.StrokeColour,
                    preset.HatchPattern.ToString(),
                    preset.HatchPitchPt,
                    preset.HatchLinePt,
                    preset.TextColour,
                    preset.StandardOutlinePt,
                    preset.ActivityHeightPt,
                    preset.MilestoneSizePt,
                    preset.DefaultLabelPosition.ToString(),
                    string.Join(
                        " ",
                        preset.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)),
                    preset.ColourCapability.ToString(),
                    preset.StrokeDash.ToString(),
                ]);
            }
        }

        // Append custom styles in their dictionary order (insertion order).
        // Normalize every row to the contract column count so a user row read
        // back from the worksheet (which may carry fewer columns than the
        // current header set) cannot overrun the write matrix.
        foreach (List<object?> customRow in customStyles.Values)
        {
            rows.Add(NormalizeRow(customRow, GanttCatalogues.StylesHeaders.Length));
        }

        return rows;
    }

    /// <summary>
    /// Copies <paramref name="source"/> into an array of exactly
    /// <paramref name="columnCount"/> columns, padding short rows with
    /// <see cref="string.Empty"/> and truncating long rows. Used to normalize
    /// user content read back from the worksheet before it is written into the
    /// rectangular <see cref="WriteOrReplaceTable"/> matrix.
    /// </summary>
    private static object?[] NormalizeRow(IReadOnlyList<object?> source, int columnCount)
    {
        object?[] result = new object?[columnCount];
        for (int column = 0; column < columnCount; column++)
        {
            result[column] = column < source.Count ? source[column] ?? string.Empty : string.Empty;
        }
        return result;
    }

    /// <summary>
    /// Builds a style row array for a custom style create, rename, or update.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A null optional member on the edit means <b>not carried</b>, so the
    /// stored value is preserved. An earlier revision rebuilt the row from
    /// hard-coded catalogue defaults and back-filled the gaps with
    /// <c>??=</c>: the fresh defaults were never null, so the back-fill never
    /// fired and an update that changed one colour silently reset the style's
    /// pattern, outline width, sizes, and label vocabulary to the catalogue
    /// defaults. This version starts from the stored row and overlays only the
    /// members the edit actually carries.
    /// </para>
    /// <para>
    /// A create has no stored row, so the catalogue defaults supply the
    /// starting point for the members the edit does not carry.
    /// </para>
    /// </remarks>
    /// <param name="custom">The validated custom style edit.</param>
    /// <param name="existingRow">The stored row for an update or rename, or null for a create.</param>
    /// <returns>The style payload row.</returns>
    private static List<object?> BuildCustomStyleRow(
        CustomStyleEdit custom,
        List<object?>? existingRow = null)
    {
        List<object?> row = existingRow is null
            ? []
            : [.. existingRow];

        // A stored row read back from an older sheet may be short; pad it so
        // the column assignments below cannot overrun it.
        while (row.Count < GanttCatalogues.StylesHeaders.Length)
        {
            row.Add(string.Empty);
        }

        row[0] = custom.NewStyleKey ?? custom.StyleKey;
        row[1] = custom.DisplayName ?? ToText(row[1]);
        row[2] = custom.FillColour;
        row[3] = custom.StrokeColour;
        row[5] = ToDoubleOrDefault(row[5], "HatchPitchPt");
        row[6] = ToDoubleOrDefault(row[6], "HatchLinePt");
        row[7] = custom.TextColour;
        row[8] = custom.StandardOutlinePt ?? ToDoubleOrDefault(row[8], "StandardOutlinePt");
        row[9] = custom.ActivityHeightPt ?? ToDoubleOrDefault(row[9], "ActivityHeightPt");
        row[10] = custom.MilestoneSizePt ?? ToDoubleOrDefault(row[10], "MilestoneSizePt");

        // The pattern, the default label position, and the permitted label
        // set are preserved from the stored row unless the edit carries them.
        row[4] = custom.HatchPattern?.ToString()
            ?? ToTextOrDefault(row[4], GanttHatchPattern.None.ToString());

        row[11] = custom.DefaultLabelPosition?.ToString()
            ?? ToTextOrDefault(row[11], GanttLabelPosition.Auto.ToString());

        row[12] = custom.AllowedLabelPositions is { } allowed
            ? string.Join(
                " ",
                allowed.Select(position => position.ToString())
                    .OrderBy(name => name, StringComparer.Ordinal))
            : ToTextOrDefault(row[12], GanttLabelPosition.Auto.ToString());

        // The capability is PRESERVED, not recomputed. It governs which
        // per-row colour overrides the row validator permits, so an update
        // must not change it as a side effect of a colour edit. A create has
        // no stored capability, so it takes the code-owned default: a custom
        // style defines the whole colour scheme, so the row's own colour
        // columns are not overridable. The hatch pattern lives in its own
        // column and needs no capability bit.
        EntityColourCapability storedCapability =
            Enum.TryParse(
                ToText(row[13]),
                ignoreCase: false,
                out EntityColourCapability parsedCapability)
                ? parsedCapability
                : EntityColourCapability.StyleDefined;
        row[13] = storedCapability.ToString();

        // The dash follows the same "not carried means preserve" rule as the
        // pattern: an edit that changes one colour must not reset the style's
        // stroke type to the built-in default.
        row[14] = custom.StrokeDash?.ToString()
            ?? ToTextOrDefault(row[14], GanttStrokeDash.Solid.ToString());

        return row;
    }

    /// <summary>
    /// Returns a cell's numeric value, or the catalogue default for the named
    /// metric when the cell is absent or unparseable.
    /// </summary>
    /// <param name="cell">The stored cell value.</param>
    /// <param name="tokenName">The catalogue metric token name.</param>
    /// <returns>The stored value, or the token's default.</returns>
    private static double ToDoubleOrDefault(object? cell, string tokenName) =>
        ToNumberOrDefault(cell, GanttCatalogues.MetricDefault(tokenName));

    /// <summary>
    /// Returns a numeric cell's value, or the supplied fallback when the cell
    /// is absent or unparseable.
    /// </summary>
    /// <param name="cell">The stored cell value.</param>
    /// <param name="fallback">The value used when the cell carries no number.</param>
    /// <returns>The cell's value, or the fallback.</returns>
    private static double ToNumberOrDefault(object? cell, double fallback) =>
        cell switch
        {
            double number => number,
            not null when double.TryParse(
                cell.ToString(),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out double parsed) => parsed,
            _ => fallback,
        };

    /// <summary>
    /// Returns a cell's text, or the supplied fallback when the cell is empty.
    /// </summary>
    /// <param name="cell">The stored cell value.</param>
    /// <param name="fallback">The value used when the cell carries no text.</param>
    /// <returns>The cell's text, or the fallback.</returns>
    private static string ToTextOrDefault(object? cell, string fallback)
    {
        string text = ToText(cell);
        return text.Length > 0 ? text : fallback;
    }

    /// <summary>
    /// Builds the themed setting rows: every catalogue setting, with the edited
    /// chart-chrome colour where the request carries one and the preserved
    /// workbook value otherwise.
    /// </summary>
    /// <param name="request">The validated theme edit.</param>
    /// <param name="preservation">The captured user content to preserve.</param>
    /// <returns>The settings payload rows.</returns>
    /// <remarks>
    /// The chart title, the time scale, the margins, and every other setting
    /// the user configured are preserved from the sheet, not regenerated from
    /// the catalogue: a colour change must not silently reset the rest of the
    /// workbook's configuration.
    /// </remarks>
    private static List<object?[]> BuildThemedSettingRows(
        ThemeEditRequest request,
        CataloguePreservation preservation)
    {
        List<object?[]> rows = [];
        foreach (GanttSettingDefinition setting in GanttCatalogues.Settings)
        {
            string? edited = null;
            _ = request.ChromeColourEdits?.TryGetValue(setting.Key, out edited);

            string value = edited
                ?? (preservation.Settings.TryGetValue(setting.Key, out string? preserved)
                    ? preserved
                    : setting.DefaultValue);

            rows.Add([setting.Key, value]);
        }

        return rows;
    }

    /// <summary>
    /// Builds themed metric rows: catalogue order, edited defaults applied.
    /// Names, minima and maxima are never edited; only the default moves.
    /// </summary>
    /// <param name="request">The validated theme edit.</param>
    /// <returns>The metric payload rows.</returns>
    private static List<object?[]> BuildThemedMetricRows(ThemeEditRequest request)
    {
        List<object?[]> rows = [];
        foreach (GanttMetricToken token in GanttCatalogues.Metrics)
        {
            double value = request.MetricEdits.TryGetValue(token.Name, out double edited)
                ? edited
                : token.DefaultValue;
            rows.Add([token.Name, value, token.Minimum, token.Maximum]);
        }

        return rows;
    }

    /// <summary>
    /// Undoes a partially applied sequence, most-recently-written first. A
    /// table that existed before is rewritten from its snapshot; a table that
    /// did not is deleted and its cell extent cleared.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="snapshots">The pre-write state, in write order.</param>
    /// <param name="written">The extent each table was written over.</param>
    /// <remarks>
    /// A failure during restoration is deliberately not swallowed: the
    /// no-mutation guarantee could not be honoured, so reporting a typed
    /// refusal would falsely promise the workbook was untouched. The host
    /// exception propagates to the command boundary instead.
    /// </remarks>
    private void RollBack(
        Worksheet config,
        List<TableSnapshot> snapshots,
        Dictionary<string, Excel.Range> written)
    {
        for (int index = snapshots.Count - 1; index >= 0; index--)
        {
            TableSnapshot snapshot = snapshots[index];
            if (snapshot.Existed)
            {
                _ = WriteOrReplaceTable(
                    config,
                    snapshot.Anchor,
                    snapshot.TableName,
                    [.. snapshot.Headers],
                    snapshot.Body);
                continue;
            }

            ListObject? created = FindTable(config, snapshot.TableName);
            created?.Delete();

            // Delete leaves the cell values behind, so the extent this call
            // wrote over must be cleared explicitly.
            if (written.TryGetValue(snapshot.TableName, out Excel.Range? extent))
            {
                ClearContents(extent);
            }
        }
    }

    /// <summary>
    /// The anchor address a contract table is written at.
    /// </summary>
    /// <param name="tableName">The contract table name.</param>
    /// <returns>The anchor cell address.</returns>
    private static string AnchorFor(string tableName) => tableName switch
    {
        GanttCatalogues.TypesTableName => GanttCatalogues.TypesAnchor,
        GanttCatalogues.StylesTableName => GanttCatalogues.StylesAnchor,
        GanttCatalogues.MetricsTableName => GanttCatalogues.MetricsAnchor,
        GanttCatalogues.SettingsTableName => GanttCatalogues.SettingsAnchor,
        _ => GanttCatalogues.ConfigAnchor,
    };

    /// <summary>
    /// Reads a table's header row through the list-column seam.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <returns>The header names, in column order.</returns>
    private List<string> ReadHeaders(ListObject table)
    {
        ListColumns columns = table.ListColumns;
        int columnCount = columns.Count;
        List<string> headers = new List<string>(columnCount);
        for (int index = 1; index <= columnCount; index++)
        {
            ListColumn column = GetColumnAt(columns, index);
            headers.Add(column.Name);
        }

        return headers;
    }

    /// <summary>
    /// Finds <c>_GanttCreatorConfig</c> by exact name. Chart sheets carry no
    /// worksheets semantics and are skipped via the worksheet cast.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <returns>The configuration worksheet, or <see langword="null"/>.</returns>
    private Worksheet? FindConfigSheet(Sheets sheets)
    {
        int count = sheets.Count;
        for (int index = 1; index <= count; index++)
        {
            object sheet = GetSheetAt(sheets, index);
            if (sheet is Worksheet worksheet
                && string.Equals(
                    worksheet.Name,
                    GanttWorkbookContract.ConfigSheetName,
                    StringComparison.Ordinal))
            {
                return worksheet;
            }
        }

        return null;
    }

    /// <summary>
    /// Captures the user content the write must preserve (ADR-0007 D4):
    /// present setting values, non-built-in style rows, and the stored
    /// workbook ID. Read-only; runs before any mutation.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>The preservation snapshot; absent tables contribute nothing.</returns>
    private CataloguePreservation? ReadPreservation(Worksheet config)
    {
        if (!TryReadKeyValues(config, GanttCatalogues.SettingsTableName, out Dictionary<string, string> settings)
            || !TryReadKeyValues(config, GanttCatalogues.ConfigTableName, out Dictionary<string, string> configValues))
        {
            return null;
        }

        List<object?[]> userStyles = ReadUserStyleRows(config);
        string? workbookId = configValues.GetValueOrDefault(GanttCatalogues.ConfigWorkbookIdKey);
        return new CataloguePreservation(settings, userStyles, workbookId);
    }

    /// <summary>
    /// Reads an existing key/value table while rejecting blank and duplicate keys.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The table to read.</param>
    /// <param name="values">The parsed key/value map when valid.</param>
    /// <returns><see langword="true"/> when the table is absent or valid.</returns>
    private bool TryReadKeyValues(
        Worksheet config,
        string tableName,
        out Dictionary<string, string> values)
    {
        values = new Dictionary<string, string>(StringComparer.Ordinal);
        ListObject? table = FindTable(config, tableName);
        if (table is null)
        {
            return true;
        }

        foreach (object?[] row in ReadBodyRows(table))
        {
            string key = ToText(row.ElementAtOrDefault(0));
            if (key.Length == 0 || !values.TryAdd(key, ToText(row.ElementAtOrDefault(1))))
            {
                values.Clear();
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Reads the rows of <c>tblGanttStyles</c> whose <c>StyleKey</c> is not
    /// a built-in preset key. Read-only.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <returns>The user rows in body order.</returns>
    private List<object?[]> ReadUserStyleRows(Worksheet config)
    {
        List<object?[]> userRows = [];
        ListObject? table = FindTable(config, GanttCatalogues.StylesTableName);
        if (table is null)
        {
            return userRows;
        }

        HashSet<string> builtInKeys = GanttCatalogues.StylePresets
            .Select(preset => preset.StyleKey)
            .ToHashSet(StringComparer.Ordinal);
        foreach (object?[] row in ReadBodyRows(table))
        {
            string styleKey = ToText(row.ElementAtOrDefault(0));
            if (styleKey.Length > 0 && !builtInKeys.Contains(styleKey))
            {
                userRows.Add(row);
            }
        }

        return userRows;
    }

    /// <summary>
    /// Bulk-reads a table body through the seams into 0-based row arrays.
    /// </summary>
    /// <param name="table">The table to read.</param>
    /// <returns>The body rows; empty when the table has no body.</returns>
    private List<object?[]> ReadBodyRows(ListObject table)
    {
        List<object?[]> rows = [];
        Excel.Range? body = GetTableBody(table);
        if (body is null)
        {
            return rows;
        }

        object? raw = GetBodyValues(body);
        foreach (object?[] row in ExcelValue2Matrix.ReadRows(raw))
        {
            rows.Add(row);
        }

        return rows;
    }


    /// <summary>
    /// Deletes the named table when it exists, then writes the header row
    /// plus the data rows as one array assignment and creates the table over
    /// the extent (one Add call). Mutates only the configuration worksheet.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="anchor">The anchor cell (e.g. <c>"A1"</c>).</param>
    /// <param name="tableName">The contract table name.</param>
    /// <param name="headers">The contract header row.</param>
    /// <param name="dataRows">The data rows (already merged with user content).</param>
    /// <param name="written">
    /// Optional dictionary that records the extent written for each table, so
    /// the rollback can clear the cells a deleted table left behind. The extent
    /// is recorded before the table is added, because the add is the call the
    /// host is most likely to refuse.
    /// </param>
    /// <returns>The extent range written.</returns>
    private Excel.Range WriteOrReplaceTable(
        Worksheet config,
        string anchor,
        string tableName,
        string[] headers,
        List<object?[]> dataRows,
        Dictionary<string, Excel.Range>? written = null)
    {
        ListObject? existing = FindTable(config, tableName);
        existing?.Delete();


        (int anchorRow, int anchorColumn) = ParseAnchor(anchor);
        int rowCount = dataRows.Count + 1;
        int columnCount = headers.Length;
        Excel.Range anchorCell = GetCellRange(config, anchorRow, anchorColumn);
        Excel.Range extent = GetResizedRange(anchorCell, rowCount, columnCount);

        // CA1814: Excel's Range.Value2 accepts only a rectangular object
        // array (a COM SAFEARRAY of VARIANT); the multidimensional form is
        // the requirement, not a style choice.
#pragma warning disable CA1814
        object[,] matrix = new object[rowCount, columnCount];
#pragma warning restore CA1814
        for (int column = 0; column < columnCount; column++)
        {
            matrix[0, column] = headers[column];
        }

        for (int row = 0; row < dataRows.Count; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                matrix[row + 1, column] = dataRows[row][column] ?? string.Empty;
            }
        }

        extent.Value2 = matrix;
        ListObjects listObjects = GetListObjects(config);

        // Record the extent before the add: if the add is what the host
        // refuses, the rollback still needs to know which cells this call
        // wrote so it can clear them.
        written?[tableName] = extent;

        ListObject table = AddTable(listObjects, extent);
        table.Name = tableName;

        // The settings table's Value column holds free text, including plot
        // dates such as "10/01/2025". Written into a General-format cell via
        // Value2, Excel re-reads that text under the host locale and stores
        // the serial it parsed (45931 = 1 Oct 2025 for a US host); the next
        // read then returns the coerced value and the UK date is lost before
        // the parser ever sees it. Formatting the column as Text AFTER the
        // table is created stores the text verbatim, which is the only form
        // the parser can parse deterministically.
        // This is a best-effort operation; if the range is not available (e.g.
        // in tests with fakes), we skip formatting rather than failing.
        if (tableName == GanttCatalogues.SettingsTableName)
        {
#pragma warning disable CA1031 // Best-effort formatting: any COM exception is swallowed
            try
            {
                ListColumns? listColumns = table.ListColumns;
                if (listColumns is not null && listColumns.Count >= 3)
                {
                    ListColumn? valueColumn = listColumns[2];
                    if (valueColumn is not null)
                    {
                        Excel.Range bodyRange = valueColumn.DataBodyRange;
                        if (bodyRange is { } br)
                        {
                            br.NumberFormat = "@";
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Best-effort: formatting failed (e.g. in tests with fakes).
                // The write still succeeds; the parser will handle legacy formats.
            }
#pragma warning restore CA1031
        }

        return extent;
    }

    /// <summary>
    /// Finds a table by exact name on the configuration worksheet.
    /// </summary>
    /// <param name="config">The configuration worksheet.</param>
    /// <param name="tableName">The contract table name.</param>
    /// <returns>The table, or <see langword="null"/>.</returns>
    private ListObject? FindTable(Worksheet config, string tableName)
    {
        ListObjects listObjects = GetListObjects(config);
        int count = listObjects.Count;
        for (int index = 1; index <= count; index++)
        {
            ListObject table = GetTableAt(listObjects, index);
            if (string.Equals(table.Name, tableName, StringComparison.Ordinal))
            {
                return table;
            }
        }

        return null;
    }

    /// <summary>
    /// Builds the <c>tblGanttTypes</c> data rows from the code-owned type
    /// catalogue. No user content exists on this table by contract.
    /// </summary>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildTypeRows() =>
    [
        .. GanttCatalogues.TypeRows
            .Select(row => new object?[]
            {
                row.TypeName,
                row.DisplayName,
                row.Kind,
                row.DateMode,
                row.DefaultStyleKey,
                row.ColourCapability,
                row.RequiresStyleKey,
                row.AllowedLabelPositions,
            }),
    ];

    /// <summary>
    /// Builds the <c>tblGanttStyles</c> data rows: the built-in presets plus
    /// the preserved user rows (ADR-0007 D4).
    /// </summary>
    /// <param name="userRows">The preserved non-built-in rows.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildStyleRows(IReadOnlyList<object?[]> userRows)
    {
        List<object?[]> rows =
        [
            .. GanttCatalogues.StylePresets
                .Select(preset => new object?[]
                {
                    preset.StyleKey,
                    preset.DisplayName,
                    preset.FillColour,
                    preset.StrokeColour,
                    preset.HatchPattern.ToString(),
                    preset.HatchPitchPt,
                    preset.HatchLinePt,
                    preset.TextColour,
                    preset.StandardOutlinePt,
                    preset.ActivityHeightPt,
                    preset.MilestoneSizePt,
                    preset.DefaultLabelPosition.ToString(),
                    string.Join(
                        " ",
                        preset.AllowedLabelPositions
                            .Select(position => position.ToString())
                            .OrderBy(name => name, StringComparer.Ordinal)),
                    preset.ColourCapability.ToString(),
                    preset.StrokeDash.ToString(),
                }),
        ];
        foreach (object?[] userRow in userRows)
        {
            rows.Add(NormalizeRow(userRow, GanttCatalogues.StylesHeaders.Length));
        }
        return rows;
    }

    /// <summary>
    /// Builds the <c>tblGanttMetrics</c> data rows. No user content exists
    /// on this table by contract.
    /// </summary>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildMetricRows() =>
    [
        .. GanttCatalogues.Metrics
            .Select(metric => new object?[]
            {
                metric.Name,
                metric.DefaultValue,
                metric.Minimum,
                metric.Maximum,
            }),
    ];

    /// <summary>
    /// Builds the <c>tblGanttSettings</c> data rows: present values win and
    /// absent keys get the schema-v2 code default, so the table always holds
    /// exactly the approved key set (ADR-0007 D2/D3 and ADR-0014). A legacy
    /// blank <c>ChartTitle</c> is replaced with the approved nonblank default;
    /// valid user title and visibility values survive regeneration. A key that
    /// is not in the approved set is not carried forward: the reader validates
    /// the exact key set, so preserving an unapproved key would leave the
    /// workbook permanently unreadable after every regeneration.
    /// </summary>
    /// <param name="existing">The preserved key/value map.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildSettingRows(IReadOnlyDictionary<string, string> existing)
    {
        List<object?[]> rows = [];
        foreach (GanttSettingDefinition setting in GanttCatalogues.Settings)
        {
            string value = existing.TryGetValue(setting.Key, out string? preserved) ? preserved : setting.DefaultValue;
            if (setting.Key == "ChartTitle" && string.IsNullOrWhiteSpace(value))
            {
                value = setting.DefaultValue;
            }

            rows.Add([setting.Key, value]);
        }

        return rows;
    }

    /// <summary>
    /// Builds the <c>tblGanttConfig</c> data rows: the schema version, the
    /// catalogue hash, the preserved-or-new workbook ID, and the add-in
    /// version last used.
    /// </summary>
    /// <param name="existingWorkbookId">The preserved workbook ID, or <see langword="null"/>.</param>
    /// <returns>The data rows.</returns>
    private static List<object?[]> BuildConfigRows(string? existingWorkbookId)
    {
        string? workbookId = existingWorkbookId;
        if (string.IsNullOrWhiteSpace(workbookId))
        {
            workbookId = Guid.NewGuid().ToString("D", CultureInfo.InvariantCulture);
        }

        return
        [
            [GanttCatalogues.ConfigSchemaVersionKey, GanttSchemaVersion.CurrentSchemaVersion.ToString(CultureInfo.InvariantCulture)],
            [GanttCatalogues.ConfigCatalogueHashKey, GanttCatalogues.ComputeCatalogueHash()],
            [GanttCatalogues.ConfigWorkbookIdKey, workbookId],
            [GanttCatalogues.ConfigAddInVersionKey, VersionInfo.SemanticVersion],
        ];
    }

    /// <summary>
    /// Converts cell text: <see langword="null"/> and Excel error values
    /// become empty; everything else is invariant text.
    /// </summary>
    /// <param name="value">The cell value.</param>
    /// <returns>The text.</returns>
    private static string ToText(object? value) =>
        value switch
        {
            null => string.Empty,
            string text => text,
            double number when double.IsNaN(number) || double.IsInfinity(number) => string.Empty,
            bool flag => flag ? "TRUE" : "FALSE",
            _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
        };

    /// <summary>
    /// Parses an A1-style anchor (e.g. <c>"A1"</c>, <c>"AB20"</c>) into
    /// one-based row and column indices, culture-invariant.
    /// </summary>
    /// <param name="anchor">The anchor address.</param>
    /// <returns>The (row, column) pair.</returns>
    private static (int Row, int Column) ParseAnchor(string anchor)
    {
        int letterCount = 0;
        while (letterCount < anchor.Length && char.IsAsciiLetterUpper(anchor[letterCount]))
        {
            letterCount++;
        }

        int column = 0;
        for (int index = 0; index < letterCount; index++)
        {
            column = (column * 26) + anchor[index] - 'A' + 1;
        }

        int row = int.Parse(
            anchor[letterCount..],
            NumberStyles.None,
            CultureInfo.InvariantCulture);
        return (row, column);
    }

    /// <summary>
    /// Returns the sheet at the one-based index. Test seam over the COM
    /// parameterised <c>Sheets.Item</c> property.
    /// </summary>
    /// <param name="sheets">The workbook's sheets.</param>
    /// <param name="index">The one-based sheet index.</param>
    /// <returns>The sheet at the index.</returns>
    internal virtual object GetSheetAt(Sheets sheets, int index) => sheets[index];

    /// <summary>
    /// Returns a worksheet's list objects. Test seam over the COM property.
    /// </summary>
    /// <param name="worksheet">The configuration worksheet.</param>
    /// <returns>The list objects.</returns>
    internal virtual ListObjects GetListObjects(Worksheet worksheet) => worksheet.ListObjects;

    /// <summary>
    /// Returns the list object at the one-based index. Test seam over the
    /// COM parameterised <c>ListObjects.Item</c> property.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="index">The one-based table index.</param>
    /// <returns>The list object at the index.</returns>
    internal virtual ListObject GetTableAt(ListObjects listObjects, int index) => listObjects[index];

    /// <summary>
    /// Returns the list column at the one-based index. Test seam over the COM
    /// parameterised <c>ListColumns.Item</c> property, used by the pre-write
    /// header capture.
    /// </summary>
    /// <param name="columns">The table's list columns.</param>
    /// <param name="index">The one-based column index.</param>
    /// <returns>The list column at the index.</returns>
    internal virtual ListColumn GetColumnAt(ListColumns columns, int index) => columns[index];

    /// <summary>
    /// Clears a range's contents without touching its formatting. Test seam
    /// over the COM <c>Range.ClearContents</c> method, used by the rollback to
    /// remove the cells a deleted table left behind.
    /// </summary>
    /// <param name="range">The range to clear.</param>
    internal virtual void ClearContents(Excel.Range range) => range.ClearContents();

    /// <summary>
    /// Adds a table over the source range with header-name behaviour. Test
    /// seam over the COM <c>ListObjects.Add</c> method.
    /// </summary>
    /// <param name="listObjects">The worksheet's list objects.</param>
    /// <param name="source">The extent range (header row plus data rows).</param>
    /// <returns>The created list object.</returns>
    internal virtual ListObject AddTable(ListObjects listObjects, object source) =>
        listObjects.Add(
            XlListObjectSourceType.xlSrcRange,
            source,
            Type.Missing,
            XlYesNoGuess.xlYes,
            Type.Missing);

    /// <summary>
    /// Returns the cell at the one-based row and column. Test seam over the
    /// COM parameterised <c>Range.Item</c> property.
    /// </summary>
    /// <param name="worksheet">The configuration worksheet.</param>
    /// <param name="row">The one-based row.</param>
    /// <param name="column">The one-based column.</param>
    /// <returns>The cell range.</returns>
    internal virtual Excel.Range GetCellRange(Worksheet worksheet, int row, int column) =>
        worksheet.Cells[row, column];

    /// <summary>
    /// Resizes a range. Test seam over the COM parameterised
    /// <c>Range.Resize</c> property.
    /// </summary>
    /// <param name="range">The anchor range.</param>
    /// <param name="rows">The row count.</param>
    /// <param name="columns">The column count.</param>
    /// <returns>The resized range.</returns>
    internal virtual Excel.Range GetResizedRange(Excel.Range range, int rows, int columns) =>
        range.Resize[rows, columns];

    /// <summary>
    /// Formats one column of a range as text. Test seam over the COM
    /// <c>Range.NumberFormat</c> property.
    /// </summary>
    /// <param name="range">The range containing the column.</param>
    /// <param name="columnIndex">The zero-based column index within <paramref name="range"/> to format.</param>
    internal virtual void FormatColumnAsText(Excel.Range range, int columnIndex) =>
        range.Columns[columnIndex + 1].NumberFormat = "@";

    /// <summary>
    /// Returns a table's body range. Test seam over the COM property.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <returns>The body range, or <see langword="null"/> when empty.</returns>
    internal virtual Excel.Range? GetTableBody(ListObject table) => table.DataBodyRange;

    /// <summary>
    /// Returns the bulk <c>Value2</c> payload of a range. Test seam so
    /// contract tests inject <c>object[,]</c> without COM.
    /// </summary>
    /// <param name="body">The range.</param>
    /// <returns>The <c>Value2</c> payload.</returns>
    internal virtual object? GetBodyValues(Excel.Range body) => body.Value2;
}
