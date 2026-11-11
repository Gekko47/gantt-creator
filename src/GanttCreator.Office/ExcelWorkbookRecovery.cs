using System.Globalization;
using GanttCreator.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace GanttCreator.Office;

/// <summary>
/// Live recovery adapter for the visible <c>tblGanttData</c> table.
/// </summary>
/// <remarks>
/// <para>
/// The adapter is split into a read-only <see cref="Inspect"/> and a mutating
/// <see cref="Recover"/>. The inspection carries the pure Core plans, so the
/// command can confirm before any host write. Recovery only fills blank header
/// cells, creates a missing table over the contract header row, or applies a
/// provable column repair; anything ambiguous refuses with no mutation.
/// </para>
/// <para>
/// COM ownership follows the repository rule: every proxy is held in a local
/// and used without chained member expressions. The parameterised COM
/// properties are exposed as <c>internal virtual</c> seams for contract tests.
/// </para>
/// </remarks>
public class ExcelWorkbookRecovery(
    object? application,
    IGanttTableReader? tableReader = null,
    IWorksheetProtectionGuard? protectionGuard = null,
    IApplicationStateScope? stateScope = null,
    IConfigCatalogueWriter? catalogueWriter = null,
    Action<string>? trace = null) : IWorkbookRecovery
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IGanttTableReader _tableReader = tableReader ?? new ExcelGanttTableReader(application);
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly IApplicationStateScope? _stateScope = stateScope;
    private readonly IConfigCatalogueWriter _catalogueWriter =
        catalogueWriter ?? new ExcelConfigCatalogueWriter(application);
    private readonly Action<string>? _trace = trace;

    /// <inheritdoc />
    public WorkbookRecoveryInspection Inspect()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        GanttTableReadOutcome read = _tableReader.Read();
        if (read.Succeeded)
        {
            if (read.ConfigurationPresent)
            {
                return new WorkbookRecoveryInspection(
                    TableReadable: true,
                    TablePresent: true,
                    ConfigurationPresent: true,
                    NeedsRecovery: false,
                    TargetSheetName: null,
                    BodyRowCount: 0,
                    HeaderPlan: null,
                    StructurePlan: null,
                    RequiresConfirmation: false,
                    ConfirmationMessage: null,
                    Refusal: null);
            }

            Excel.Sheets successSheets = workbook.Sheets;
            string? tableSheet = TryFindTable(successSheets, out Excel.Worksheet? gantt, out _)
                ? gantt?.Name
                : null;
            return new WorkbookRecoveryInspection(
                TableReadable: true,
                TablePresent: true,
                ConfigurationPresent: false,
                NeedsRecovery: true,
                TargetSheetName: tableSheet,
                BodyRowCount: 0,
                HeaderPlan: null,
                StructurePlan: null,
                RequiresConfirmation: false,
                ConfirmationMessage: null,
                Refusal: null)
            {
                NeedsConfigurationRecovery = true,
            };
        }

        if (read.Refusal == GanttTableReadRefusalReason.NoActiveWorkbook)
        {
            return Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        if (read.Refusal == GanttTableReadRefusalReason.DateSystemUnsupported)
        {
            return Refused(WorkbookRecoveryRefusalReason.DateSystemUnsupported);
        }

        Excel.Sheets sheets = workbook.Sheets;
        return read.Refusal == GanttTableReadRefusalReason.TableUnreadable
            ? InspectStructureRepair(sheets, read)
            : InspectRebuild(sheets, read);
    }

    private WorkbookRecoveryInspection InspectStructureRepair(
        Excel.Sheets sheets,
        GanttTableReadOutcome read)
    {
        if (!TryFindTable(sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetMissing, read.ConfigurationPresent);
        }

        Excel.ListColumns columns = GetTableColumns(table);
        int count = GetColumnCount(columns);
        List<string> columnHeaders = new(count);
        for (int index = 1; index <= count; index++)
        {
            columnHeaders.Add(GetColumnName(GetColumnAt(columns, index)));
        }

        GanttTableStructureRepairPlan plan = GanttTableStructureRepair.Plan(columnHeaders);
        bool canRepair = plan.Succeeded;
        return new WorkbookRecoveryInspection(
            TableReadable: false,
            TablePresent: true,
            read.ConfigurationPresent,
            NeedsRecovery: canRepair,
            worksheet.Name,
            BodyRowCount: 0,
            HeaderPlan: null,
            plan,
            RequiresConfirmation: false,
            ConfirmationMessage: null,
            canRepair ? null : WorkbookRecoveryRefusalReason.TableStructureRefused)
        {
            NeedsConfigurationRecovery = !read.ConfigurationPresent,
        };
    }

    private WorkbookRecoveryInspection InspectRebuild(
        Excel.Sheets sheets,
        GanttTableReadOutcome read)
    {
        List<Excel.Worksheet> labelled = FindLabelledSheets(sheets);
        if (labelled.Count == 0)
        {
            return !read.ConfigurationPresent
                ? Refused(WorkbookRecoveryRefusalReason.TargetMissing)
                : new WorkbookRecoveryInspection(
                    TableReadable: false,
                    TablePresent: false,
                    ConfigurationPresent: true,
                    NeedsRecovery: true,
                    TargetSheetName: null,
                    BodyRowCount: 0,
                    HeaderPlan: null,
                    StructurePlan: null,
                    RequiresConfirmation: true,
                    ConfirmationMessage:
                        "The Gantt worksheet is missing. Fix can create a new empty one from the stored configuration. Your schedule rows are not restored. Continue?",
                    Refusal: null)
                {
                    NeedsGanttSheet = true,
                };
        }

        if (labelled.Count > 1)
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetAmbiguous, read.ConfigurationPresent);
        }

        Excel.Worksheet target = labelled[0];
        if (HasAnyTable(target))
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetHasForeignTable, read.ConfigurationPresent);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(target);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected,
                read.ConfigurationPresent);
        }

        GanttTableHeaderAlignmentPlan? headerPlan = BuildHeaderPlan(target);
        if (headerPlan is null || !headerPlan.Succeeded)
        {
            return new WorkbookRecoveryInspection(
                TableReadable: false,
                TablePresent: false,
                read.ConfigurationPresent,
                NeedsRecovery: false,
                target.Name,
                BodyRowCount: 0,
                headerPlan,
                StructurePlan: null,
                RequiresConfirmation: false,
                ConfirmationMessage: null,
                WorkbookRecoveryRefusalReason.HeaderMisaligned);
        }

        int bodyRowCount = CountBodyRows(target);
        bool confirm = bodyRowCount > 0;
        string? message = confirm
            ? string.Create(
                CultureInfo.InvariantCulture,
                $"Fix will rebuild tblGanttData on '{target.Name}' and adopt {bodyRowCount} existing row(s). Continue?")
            : null;

        return new WorkbookRecoveryInspection(
            TableReadable: false,
            TablePresent: false,
            read.ConfigurationPresent,
            NeedsRecovery: true,
            target.Name,
            bodyRowCount,
            headerPlan,
            StructurePlan: null,
            confirm,
            message,
            Refusal: null)
        {
            NeedsConfigurationRecovery = !read.ConfigurationPresent,
        };
    }

    /// <inheritdoc />
    public WorkbookRecoveryOutcome Recover(WorkbookRecoveryInspection inspection, bool confirmed)
    {
        ArgumentNullException.ThrowIfNull(inspection);

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookRecoveryOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected);
        }

        if (inspection.Refusal is { } refusal)
        {
            return WorkbookRecoveryOutcome.Refused(refusal);
        }

        if (!inspection.NeedsRecovery)
        {
            return WorkbookRecoveryOutcome.NoChange();
        }

        if (inspection.RequiresConfirmation && !confirmed)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.ConfirmationDeclined);
        }

        using IApplicationStateScope scope = _stateScope ?? NullApplicationStateScope.Instance;
        scope.SuppressScreenUpdating();
        scope.SuppressEvents();
        scope.SuppressAlerts();
        scope.CaptureSelection();

        if (inspection.NeedsGanttSheet)
        {
            return ApplyCreateGanttSheet();
        }

        bool configurationCreated = false;
        if (inspection.NeedsConfigurationRecovery)
        {
            WorkbookRecoveryRefusalReason? configRefusal = CreateConfigurationSheet(inspection);
            if (configRefusal is not null)
            {
                return WorkbookRecoveryOutcome.Refused(configRefusal.Value);
            }

            configurationCreated = true;
            if (inspection.TableReadable && inspection.StructurePlan is null)
            {
                return new WorkbookRecoveryOutcome(
                    Changed: true,
                    TableRebuilt: false,
                    TableStructureRepaired: false,
                    AdoptedRowCount: 0,
                    Refusal: null)
                {
                    ConfigurationCreated = true,
                };
            }
        }

        WorkbookRecoveryOutcome outcome = inspection.StructurePlan is { } structure
            ? ApplyStructureRepair(structure)
            : ApplyRebuild(inspection);
        if (outcome.Refusal is not null && configurationCreated)
        {
            DeleteCreatedConfigurationSheet();
        }

        return outcome with { ConfigurationCreated = configurationCreated };
    }

    private WorkbookRecoveryOutcome ApplyRebuild(WorkbookRecoveryInspection inspection)
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindLabelledSheet(sheets, inspection.TargetSheetName, out Excel.Worksheet? target) || target is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TargetMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(target);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookRecoveryOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected);
        }

        GanttTableHeaderAlignmentPlan headerPlan =
            inspection.HeaderPlan
            ?? new GanttTableHeaderAlignmentPlan(GanttTableHeaderAlignmentOutcome.Aligned, []);
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        string?[] previousHeaders = ReadHeaderValues(target);

        Excel.ListObject? created = null;
        try
        {
            foreach (int index in headerPlan.CellsToWrite)
            {
                SetHeaderValue(target, index + 1, schema[index].Name);
            }

            Excel.Range range = GetTableRange(
                target,
                GanttSheetLayout.HeaderRowIndex,
                inspection.BodyRowCount,
                schema.Count);
            created = AddTable(target, range);
            PresentTable(created);
            ApplyColumnPresentation(created);

            GanttTableReadOutcome verify = _tableReader.Read();
            if (!verify.Succeeded)
            {
                RollBackRebuild(target, created, previousHeaders);
                return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.VerificationFailed);
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery rebuilt table={GanttTableSchema.TableName} sheet={target.Name} adopted={inspection.BodyRowCount}"));

            return new WorkbookRecoveryOutcome(
                Changed: true,
                TableRebuilt: true,
                TableStructureRepaired: false,
                inspection.BodyRowCount,
                Refusal: null);
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery rebuild refused " + exception.Message);
            RollBackRebuild(target, created, previousHeaders);
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TableCreateFailed);
        }
    }

    private WorkbookRecoveryOutcome ApplyStructureRepair(GanttTableStructureRepairPlan plan)
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindTable(sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TargetMissing);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(worksheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookRecoveryOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected);
        }

        Excel.ListColumns columns = GetTableColumns(table);
        int originalCount = GetColumnCount(columns);
        string[] originalNames = new string[originalCount];
        for (int index = 0; index < originalCount; index++)
        {
            originalNames[index] = GetColumnName(GetColumnAt(columns, index + 1));
        }

        List<int> insertedPositions = [];
        try
        {
            foreach (GanttTableColumnRepair repair in plan.Repairs)
            {
                if (repair.Kind == GanttTableColumnRepairKind.Rename)
                {
                    Excel.ListColumn column = GetColumnAt(columns, repair.SchemaIndex + 1);
                    SetColumnName(column, repair.Name);
                }
                else
                {
                    Excel.ListColumn added = AddColumnAt(columns, repair.SchemaIndex + 1);
                    SetColumnName(added, repair.Name);
                    insertedPositions.Add(repair.SchemaIndex + 1);
                }
            }

            GanttTableReadOutcome verify = _tableReader.Read();
            if (!verify.Succeeded)
            {
                RollBackStructureRepair(columns, originalNames, insertedPositions);
                return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.VerificationFailed);
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery repaired table structure sheet={worksheet.Name} changes={plan.Repairs.Count}"));

            return new WorkbookRecoveryOutcome(
                Changed: true,
                TableRebuilt: false,
                TableStructureRepaired: true,
                AdoptedRowCount: 0,
                Refusal: null);
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery structure refused " + exception.Message);
            RollBackStructureRepair(columns, originalNames, insertedPositions);
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.HeaderWriteFailed);
        }
    }

    private WorkbookRecoveryOutcome ApplyCreateGanttSheet()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookRecoveryOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected);
        }

        Excel.Sheets sheets = workbook.Sheets;
        Excel.Worksheet? created = null;
        try
        {
            created = AddSheetAfter(sheets, null);
            SetSheetName(created, GanttWorkbookContract.GanttSheetLabel);

            IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
            for (int index = 0; index < schema.Count; index++)
            {
                SetHeaderValue(created, index + 1, schema[index].Name);
            }

            Excel.Range range = GetTableRange(created, GanttSheetLayout.HeaderRowIndex, 0, schema.Count);
            Excel.ListObject table = AddTable(created, range);
            PresentTable(table);
            ApplyColumnPresentation(table);

            GanttTableReadOutcome verify = _tableReader.Read();
            if (!verify.Succeeded)
            {
                DeleteSheet(created);
                return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.VerificationFailed);
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery created Gantt sheet={created.Name}"));
            return new WorkbookRecoveryOutcome(
                Changed: true,
                TableRebuilt: true,
                TableStructureRepaired: false,
                AdoptedRowCount: 0,
                Refusal: null)
            {
                GanttSheetCreated = true,
            };
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery Gantt sheet create refused " + exception.Message);
            if (created is not null)
            {
                DeleteSheet(created);
            }

            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.GanttSheetCreateFailed);
        }
    }

    private void DeleteSheet(Excel.Worksheet sheet)
    {
        Excel.Application? application = _application;
        if (application is null)
        {
            return;
        }

        bool original = application.DisplayAlerts;
        try
        {
            application.DisplayAlerts = false;
            sheet.Delete();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort rollback; the original refusal is the reported outcome.
        }
        finally
        {
            application.DisplayAlerts = original;
        }
    }

    private WorkbookRecoveryRefusalReason? CreateConfigurationSheet(
        WorkbookRecoveryInspection inspection)
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return WorkbookRecoveryRefusalReason.NoActiveWorkbook;
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (FindConfigurationSheet(sheets) is not null)
        {
            return null;
        }

        ProtectionGuardOutcome protection = _protectionGuard.Query();
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return protection == ProtectionGuardOutcome.NoActiveWorkbook
                ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                : WorkbookRecoveryRefusalReason.TargetProtected;
        }

        Excel.Worksheet? anchor = null;
        if (!string.IsNullOrWhiteSpace(inspection.TargetSheetName))
        {
            _ = TryFindLabelledSheet(sheets, inspection.TargetSheetName, out anchor);
        }

        Excel.Worksheet? created = null;
        try
        {
            created = AddSheetAfter(sheets, anchor);
            SetSheetName(created, GanttWorkbookContract.ConfigSheetName);
            SetSheetVisibility(created, Excel.XlSheetVisibility.xlSheetVeryHidden);

            ConfigWriteOutcome write = _catalogueWriter.Write();
            if (!write.Succeeded)
            {
                DeleteCreatedConfigurationSheet();
                return WorkbookRecoveryRefusalReason.ConfigurationWriteFailed;
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery created configuration sheet={created.Name}"));
            return null;
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery configuration create refused " + exception.Message);
            if (created is not null)
            {
                DeleteCreatedConfigurationSheet();
            }

            return WorkbookRecoveryRefusalReason.ConfigurationCreateFailed;
        }
    }

    private void DeleteCreatedConfigurationSheet()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return;
        }

        Excel.Sheets sheets = workbook.Sheets;
        Excel.Worksheet? config = FindConfigurationSheet(sheets);
        if (config is null)
        {
            return;
        }

        bool original = application.DisplayAlerts;
        try
        {
            application.DisplayAlerts = false;
            if (config.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden)
            {
                config.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            }

            config.Delete();
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort rollback; the original refusal is the reported outcome.
        }
        finally
        {
            application.DisplayAlerts = original;
        }
    }

    private Excel.Worksheet? FindConfigurationSheet(Excel.Sheets sheets)
    {
        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is Excel.Worksheet worksheet
                && string.Equals(
                    worksheet.Name,
                    GanttWorkbookContract.ConfigSheetName,
                    StringComparison.OrdinalIgnoreCase))
            {
                return worksheet;
            }
        }

        return null;
    }

    private void RollBackRebuild(
        Excel.Worksheet target,
        Excel.ListObject? created,
        string?[] previousHeaders)
    {
        try
        {
            if (created is not null)
            {
                DeleteTable(created);
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort rollback; the original refusal is the reported outcome.
        }

        try
        {
            for (int index = 0; index < previousHeaders.Length; index++)
            {
                SetHeaderValueOrBlank(target, index + 1, previousHeaders[index]);
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort rollback.
        }
    }

    private void RollBackStructureRepair(
        Excel.ListColumns columns,
        string[] originalNames,
        List<int> insertedPositions)
    {
        try
        {
            for (int index = 0; index < originalNames.Length && index + 1 <= GetColumnCount(columns); index++)
            {
                Excel.ListColumn column = GetColumnAt(columns, index + 1);
                if (!string.Equals(GetColumnName(column), originalNames[index], StringComparison.Ordinal))
                {
                    SetColumnName(column, originalNames[index]);
                }
            }

            for (int index = insertedPositions.Count - 1; index >= 0; index--)
            {
                int position = insertedPositions[index];
                if (position <= GetColumnCount(columns))
                {
                    DeleteColumn(GetColumnAt(columns, position));
                }
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort rollback.
        }
    }

    private static WorkbookRecoveryInspection Refused(
        WorkbookRecoveryRefusalReason refusal,
        bool configurationPresent = false) =>
        new(
            TableReadable: false,
            TablePresent: false,
            configurationPresent,
            NeedsRecovery: false,
            TargetSheetName: null,
            BodyRowCount: 0,
            HeaderPlan: null,
            StructurePlan: null,
            RequiresConfirmation: false,
            ConfirmationMessage: null,
            refusal);

    private GanttTableHeaderAlignmentPlan? BuildHeaderPlan(Excel.Worksheet target)
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        string?[] values = new string?[schema.Count];
        for (int index = 0; index < schema.Count; index++)
        {
            object? value = GetHeaderValue(target, index + 1);
            values[index] = value is null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        return GanttTableHeaderAlignment.Plan(values);
    }

    private string?[] ReadHeaderValues(Excel.Worksheet target)
    {
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        string?[] values = new string?[schema.Count];
        for (int index = 0; index < schema.Count; index++)
        {
            object? value = GetHeaderValue(target, index + 1);
            values[index] = value is null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        return values;
    }

    private int CountBodyRows(Excel.Worksheet target)
    {
        int firstBodyRow = GanttSheetLayout.HeaderRowIndex + 1;
        int lastUsedRow = GetUsedRangeLastRow(target);
        if (lastUsedRow < firstBodyRow)
        {
            return 0;
        }

        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        Excel.Range body = GetRange(target, firstBodyRow, 1, lastUsedRow, schema.Count);
        List<object?[]> rows = ExcelValue2Matrix.ReadRows(GetRangeValues(body));
        int count = rows.Count;
        while (count > 0 && rows[count - 1].All(IsBlank))
        {
            count--;
        }

        return count;
    }

    private static bool IsBlank(object? value) =>
        value is null || (value is string text && string.IsNullOrWhiteSpace(text));

    private List<Excel.Worksheet> FindLabelledSheets(Excel.Sheets sheets)
    {
        List<Excel.Worksheet> labelled = [];
        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is Excel.Worksheet worksheet
                && string.Equals(
                    worksheet.Name,
                    GanttWorkbookContract.GanttSheetLabel,
                    StringComparison.OrdinalIgnoreCase))
            {
                labelled.Add(worksheet);
            }
        }

        return labelled;
    }

    private bool TryFindLabelledSheet(
        Excel.Sheets sheets,
        string? name,
        out Excel.Worksheet? worksheet)
    {
        worksheet = null;
        if (string.IsNullOrWhiteSpace(name))
        {
            return false;
        }

        int count = GetSheetCount(sheets);
        for (int index = 1; index <= count; index++)
        {
            if (GetSheetAt(sheets, index) is Excel.Worksheet candidate
                && string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                worksheet = candidate;
                return true;
            }
        }

        return false;
    }

    private bool TryFindTable(
        Excel.Sheets sheets,
        out Excel.Worksheet? worksheet,
        out Excel.ListObject? table)
    {
        worksheet = null;
        table = null;
        int sheetCount = GetSheetCount(sheets);
        for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            if (GetSheetAt(sheets, sheetIndex) is not Excel.Worksheet candidate)
            {
                continue;
            }

            Excel.ListObjects objects = candidate.ListObjects;
            int objectCount = GetTableCount(objects);
            for (int objectIndex = 1; objectIndex <= objectCount; objectIndex++)
            {
                Excel.ListObject candidateTable = GetTableAt(objects, objectIndex);
                if (string.Equals(
                    GetTableName(candidateTable),
                    GanttTableSchema.TableName,
                    StringComparison.OrdinalIgnoreCase))
                {
                    worksheet = candidate;
                    table = candidateTable;
                    return true;
                }
            }
        }

        return false;
    }

    private bool HasAnyTable(Excel.Worksheet worksheet)
    {
        Excel.ListObjects objects = worksheet.ListObjects;
        return GetTableCount(objects) > 0;
    }

    private void PresentTable(Excel.ListObject table)
    {
        SetTableName(table, GanttTableSchema.TableName);
        SetShowAutoFilter(table, false);
        SetShowTableStyleRowStripes(table, false);
        SetShowTableStyleColumnStripes(table, false);
    }

    private void ApplyColumnPresentation(Excel.ListObject table)
    {
        Excel.ListColumns columns = GetTableColumns(table);
        IReadOnlyList<GanttTableColumn> schema = GanttTableSchema.Default.Columns;
        for (int index = 1; index <= schema.Count; index++)
        {
            Excel.ListColumn column = GetColumnAt(columns, index);
            Excel.Range columnRange = GetColumnRange(column);
            SetColumnLocked(columnRange, schema[index - 1].IsLocked);
            Excel.Range entireColumn = GetEntireColumn(columnRange);
            bool wanted = schema[index - 1].IsHidden;
            if (ReadHiddenFlag(entireColumn) != wanted)
            {
                SetColumnHidden(entireColumn, wanted);
            }
        }
    }

    private static bool ReadHiddenFlag(Excel.Range range) => range.Hidden switch
    {
        bool flag => flag,
        null => false,
        _ => bool.TryParse(
            Convert.ToString(range.Hidden, CultureInfo.InvariantCulture),
            out bool parsed) && parsed,
    };

    // ---- Test seams ----

    internal virtual int GetSheetCount(Excel.Sheets sheets) => sheets.Count;

    internal virtual object GetSheetAt(Excel.Sheets sheets, int index) => sheets[index];

    internal virtual int GetTableCount(Excel.ListObjects objects) => objects.Count;

    internal virtual Excel.ListObject GetTableAt(Excel.ListObjects objects, int index) => objects[index];

    internal virtual string GetTableName(Excel.ListObject table) => table.Name;

    internal virtual void SetTableName(Excel.ListObject table, string name) => table.Name = name;

    internal virtual void SetShowAutoFilter(Excel.ListObject table, bool value) => table.ShowAutoFilter = value;

    internal virtual void SetShowTableStyleRowStripes(Excel.ListObject table, bool value) =>
        table.ShowTableStyleRowStripes = value;

    internal virtual void SetShowTableStyleColumnStripes(Excel.ListObject table, bool value) =>
        table.ShowTableStyleColumnStripes = value;

    internal virtual Excel.ListColumns GetTableColumns(Excel.ListObject table) => table.ListColumns;

    internal virtual int GetColumnCount(Excel.ListColumns columns) => columns.Count;

    internal virtual Excel.ListColumn GetColumnAt(Excel.ListColumns columns, int index) => columns[index];

    internal virtual Excel.ListColumn AddColumnAt(Excel.ListColumns columns, int position) =>
        columns.Add(position);

    internal virtual string GetColumnName(Excel.ListColumn column) => column.Name;

    internal virtual void SetColumnName(Excel.ListColumn column, string name) => column.Name = name;

    internal virtual Excel.Range GetColumnRange(Excel.ListColumn column) => column.Range;

    internal virtual Excel.Range GetEntireColumn(Excel.Range range) => range.EntireColumn;

    internal virtual void SetColumnLocked(Excel.Range range, bool value) => range.Locked = value;

    internal virtual void SetColumnHidden(Excel.Range range, bool value) => range.Hidden = value;

    internal virtual object? GetHeaderValue(Excel.Worksheet target, int columnIndex) =>
        target.Cells[GanttSheetLayout.HeaderRowIndex, columnIndex].Value2;

    internal virtual void SetHeaderValue(Excel.Worksheet target, int columnIndex, string? value) =>
        target.Cells[GanttSheetLayout.HeaderRowIndex, columnIndex].Value2 = value;

    private void SetHeaderValueOrBlank(Excel.Worksheet target, int columnIndex, string? value)
    {
        object? current = GetHeaderValue(target, columnIndex);
        if (!string.Equals(
                Convert.ToString(current, CultureInfo.InvariantCulture),
                value,
                StringComparison.Ordinal))
        {
            SetHeaderValue(target, columnIndex, value);
        }
    }

    internal virtual int GetUsedRangeLastRow(Excel.Worksheet target)
    {
        Excel.Range used = target.UsedRange;
        return used.Row + used.Rows.Count - 1;
    }

    internal virtual Excel.Range GetRange(
        Excel.Worksheet target,
        int firstRow,
        int firstColumn,
        int lastRow,
        int lastColumn) =>
        target.Range[target.Cells[firstRow, firstColumn], target.Cells[lastRow, lastColumn]];

    internal virtual object? GetRangeValues(Excel.Range range) => range.Value2;

    internal virtual Excel.Range GetTableRange(
        Excel.Worksheet target,
        int headerRow,
        int bodyRowCount,
        int columnCount) =>
        target.Range[
            target.Cells[headerRow, 1],
            target.Cells[headerRow + bodyRowCount, columnCount]];

    internal virtual Excel.ListObject AddTable(Excel.Worksheet target, Excel.Range range)
    {
        Excel.ListObjects objects = target.ListObjects;
        return objects.Add(
            Excel.XlListObjectSourceType.xlSrcRange,
            range,
            Type.Missing,
            Excel.XlYesNoGuess.xlYes,
            Type.Missing);
    }

    internal virtual void DeleteTable(Excel.ListObject table)
    {
        if (_application is not null)
        {
            bool original = _application.DisplayAlerts;
            try
            {
                _application.DisplayAlerts = false;
                table.Delete();
            }
            finally
            {
                _application.DisplayAlerts = original;
            }
        }
        else
        {
            table.Delete();
        }
    }

    internal virtual void DeleteColumn(Excel.ListColumn column) => column.Delete();

    internal virtual Excel.Worksheet AddSheetAfter(Excel.Sheets sheets, Excel.Worksheet? after) =>
        after is null ? (Excel.Worksheet)sheets.Add() : (Excel.Worksheet)sheets.Add(After: after);

    internal virtual void SetSheetName(Excel.Worksheet sheet, string name) => sheet.Name = name;

    internal virtual void SetSheetVisibility(Excel.Worksheet sheet, Excel.XlSheetVisibility visibility) =>
        sheet.Visible = visibility;
}
