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
    Action<string>? trace = null,
    ExcelConfigSheetInspector? configSheetInspector = null) : IWorkbookRecovery
{
    private readonly Excel.Application? _application = application as Excel.Application;
    private readonly IGanttTableReader _tableReader = tableReader ?? new ExcelGanttTableReader(application);
    private readonly IWorksheetProtectionGuard _protectionGuard =
        protectionGuard ?? new ExcelWorksheetProtectionGuard(application);
    private readonly IApplicationStateScope? _stateScope = stateScope;
    private readonly IConfigCatalogueWriter _catalogueWriter =
        catalogueWriter ?? new ExcelConfigCatalogueWriter(application);
    private readonly Action<string>? _trace = trace;
    private readonly ExcelConfigSheetInspector _configSheets =
        configSheetInspector ?? new ExcelConfigSheetInspector();

    /// <inheritdoc />
    public WorkbookRecoveryInspection Inspect()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        ConfigSheetLocation configSheet = _configSheets.Locate(sheets);
        if (RefusalForConfigState(configSheet.State) is { } configRefusal)
        {
            return Refused(configRefusal);
        }

        bool configurationNeedsRename = configSheet.NeedsRename;
        bool configurationMissing = configSheet.State == ConfigSheetState.Missing;
        bool configurationPresent =
            configSheet.State is ConfigSheetState.Present or ConfigSheetState.Renamed;

        GanttTableReadOutcome read = _tableReader.Read();
        if (read.Succeeded)
        {
            if (read.ConfigurationPresent && !configurationNeedsRename)
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

            string? tableSheet = null;
            if (TryFindTable(sheets, out Excel.Worksheet? gantt, out _) && gantt is not null)
            {
                tableSheet = GetSheetName(gantt);
            }

            return new WorkbookRecoveryInspection(
                TableReadable: true,
                TablePresent: true,
                configurationPresent,
                NeedsRecovery: true,
                TargetSheetName: tableSheet,
                BodyRowCount: 0,
                HeaderPlan: null,
                StructurePlan: null,
                RequiresConfirmation: false,
                ConfirmationMessage: null,
                Refusal: null)
            {
                NeedsConfigurationRecovery = configurationMissing,
                NeedsConfigurationRename = configurationNeedsRename,
            };
        }

        return read.Refusal switch
        {
            GanttTableReadRefusalReason.NoActiveWorkbook =>
                Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook),
            GanttTableReadRefusalReason.DateSystemUnsupported =>
                Refused(WorkbookRecoveryRefusalReason.DateSystemUnsupported),
            GanttTableReadRefusalReason.TableUnreadable =>
                InspectStructureRepair(
                    sheets,
                    configurationPresent,
                    configurationMissing,
                    configurationNeedsRename),
            GanttTableReadRefusalReason.TableMissing =>
                InspectRebuild(
                    sheets,
                    configurationPresent,
                    configurationMissing,
                    configurationNeedsRename),
            null => InspectRebuild(
                sheets,
                configurationPresent,
                configurationMissing,
                configurationNeedsRename),
            _ => InspectRebuild(
                sheets,
                configurationPresent,
                configurationMissing,
                configurationNeedsRename),
        };
    }

    private WorkbookRecoveryInspection InspectStructureRepair(
        Excel.Sheets sheets,
        bool configurationPresent,
        bool configurationMissing,
        bool configurationNeedsRename)
    {
        if (!TryFindTable(sheets, out Excel.Worksheet? worksheet, out Excel.ListObject? table)
            || worksheet is null
            || table is null)
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetMissing, configurationPresent);
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
            configurationPresent,
            NeedsRecovery: canRepair,
            GetSheetName(worksheet),
            BodyRowCount: 0,
            HeaderPlan: null,
            plan,
            RequiresConfirmation: false,
            ConfirmationMessage: null,
            canRepair ? null : WorkbookRecoveryRefusalReason.TableStructureRefused)
        {
            NeedsConfigurationRecovery = configurationMissing,
            NeedsConfigurationRename = configurationNeedsRename,
        };
    }

    private WorkbookRecoveryInspection InspectRebuild(
        Excel.Sheets sheets,
        bool configurationPresent,
        bool configurationMissing,
        bool configurationNeedsRename)
    {
        List<Excel.Worksheet> labelled = FindLabelledSheets(sheets);
        if (labelled.Count == 0)
        {
            return !configurationPresent
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
                    NeedsConfigurationRename = configurationNeedsRename,
                };
        }

        if (labelled.Count > 1)
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetAmbiguous, configurationPresent);
        }

        Excel.Worksheet target = labelled[0];
        Excel.ListObject? foreignTable = GetSingleTable(target);
        if (foreignTable is not null)
        {
            GanttTableHeaderAlignmentPlan? identity = BuildHeaderPlan(target);
            bool provesGanttChart =
                identity is { Outcome: GanttTableHeaderAlignmentOutcome.Aligned }
                && HasPlotAnchor(target);
            return provesGanttChart
                ? new WorkbookRecoveryInspection(
                    TableReadable: false,
                    TablePresent: true,
                    configurationPresent,
                    NeedsRecovery: true,
                    GetSheetName(target),
                    BodyRowCount: 0,
                    HeaderPlan: null,
                    StructurePlan: null,
                    RequiresConfirmation: false,
                    ConfirmationMessage: null,
                    Refusal: null)
                {
                    NeedsConfigurationRecovery = configurationMissing,
                    NeedsConfigurationRename = configurationNeedsRename,
                    NeedsTableRename = true,
                    TargetTableName = GetTableName(foreignTable),
                }
                : Refused(WorkbookRecoveryRefusalReason.TargetHasForeignTable, configurationPresent);
        }

        if (HasAnyTable(target))
        {
            return Refused(WorkbookRecoveryRefusalReason.TargetHasForeignTable, configurationPresent);
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(target);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected,
                configurationPresent);
        }

        GanttTableHeaderAlignmentPlan? headerPlan = BuildHeaderPlan(target);
        if (headerPlan is null || !headerPlan.Succeeded)
        {
            return new WorkbookRecoveryInspection(
                TableReadable: false,
                TablePresent: false,
                configurationPresent,
                NeedsRecovery: false,
                GetSheetName(target),
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
            configurationPresent,
            NeedsRecovery: true,
            GetSheetName(target),
            bodyRowCount,
            headerPlan,
            StructurePlan: null,
            confirm,
            message,
            Refusal: null)
        {
            NeedsConfigurationRecovery = configurationMissing,
            NeedsConfigurationRename = configurationNeedsRename,
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

        bool configurationCreated = false;
        bool configurationRenamed = false;
        if (inspection.NeedsConfigurationRename || inspection.NeedsConfigurationRecovery)
        {
            WorkbookRecoveryRefusalReason? configRefusal = inspection.NeedsConfigurationRename
                ? RenameConfigurationSheet()
                : CreateConfigurationSheet(inspection);
            if (configRefusal is not null)
            {
                return WorkbookRecoveryOutcome.Refused(configRefusal.Value);
            }

            configurationCreated = inspection.NeedsConfigurationRecovery;
            configurationRenamed = inspection.NeedsConfigurationRename;
        }

        if (inspection.NeedsGanttSheet)
        {
            return ApplyCreateGanttSheet() with
            {
                ConfigurationCreated = configurationCreated,
                ConfigurationRenamed = configurationRenamed,
            };
        }

        if (inspection.NeedsTableRename)
        {
            return ApplyTableRename(inspection) with
            {
                ConfigurationCreated = configurationCreated,
                ConfigurationRenamed = configurationRenamed,
            };
        }

        if (inspection.TableReadable && inspection.StructurePlan is null)
        {
            return new WorkbookRecoveryOutcome(
                Changed: true,
                TableRebuilt: false,
                TableStructureRepaired: false,
                AdoptedRowCount: 0,
                Refusal: null)
            {
                ConfigurationCreated = configurationCreated,
                ConfigurationRenamed = configurationRenamed,
            };
        }

        WorkbookRecoveryOutcome outcome = inspection.StructurePlan is { } structure
            ? ApplyStructureRepair(structure)
            : ApplyRebuild(inspection);
        if (outcome.Refusal is not null && configurationCreated)
        {
            DeleteCreatedConfigurationSheet();
        }

        return outcome with
        {
            ConfigurationCreated = configurationCreated,
            ConfigurationRenamed = configurationRenamed,
        };
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
        if (!TryFindSheetByName(sheets, inspection.TargetSheetName, out Excel.Worksheet? target) || target is null)
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

    private WorkbookRecoveryOutcome ApplyTableRename(WorkbookRecoveryInspection inspection)
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null || string.IsNullOrWhiteSpace(inspection.TargetSheetName))
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.NoActiveWorkbook);
        }

        Excel.Sheets sheets = workbook.Sheets;
        if (!TryFindSheetByName(sheets, inspection.TargetSheetName, out Excel.Worksheet? target)
            || target is null)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TargetMissing);
        }

        Excel.ListObjects objects = target.ListObjects;
        if (GetTableCount(objects) != 1)
        {
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TargetHasForeignTable);
        }

        Excel.ListObject table = GetTableAt(objects, 1);
        string previousName = GetTableName(table);
        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(target);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return WorkbookRecoveryOutcome.Refused(
                protection == ProtectionGuardOutcome.NoActiveWorkbook
                    ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                    : WorkbookRecoveryRefusalReason.TargetProtected);
        }

        try
        {
            SetTableName(table, GanttTableSchema.TableName);
            SetShowAutoFilter(table, false);
            if (!_tableReader.Read().Succeeded)
            {
                SetTableName(table, previousName);
                return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.VerificationFailed);
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery renamed table={previousName} to={GanttTableSchema.TableName} sheet={GetSheetName(target)}"));
            return new WorkbookRecoveryOutcome(
                Changed: true,
                TableRebuilt: false,
                TableStructureRepaired: false,
                AdoptedRowCount: 0,
                Refusal: null)
            {
                TableRenamed = true,
            };
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery table rename refused " + exception.Message);
            return WorkbookRecoveryOutcome.Refused(WorkbookRecoveryRefusalReason.TableRenameFailed);
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
            _ = TryFindSheetByName(sheets, inspection.TargetSheetName, out anchor);
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

    private WorkbookRecoveryRefusalReason? RenameConfigurationSheet()
    {
        Excel.Application? application = _application;
        Excel.Workbook? workbook = application?.ActiveWorkbook;
        if (application is null || workbook is null)
        {
            return WorkbookRecoveryRefusalReason.NoActiveWorkbook;
        }

        Excel.Sheets sheets = workbook.Sheets;
        ConfigSheetLocation location = _configSheets.Locate(sheets);
        if (!location.NeedsRename
            || string.IsNullOrWhiteSpace(location.SheetName)
            || !TryFindSheetByName(sheets, location.SheetName, out Excel.Worksheet? sheet)
            || sheet is null)
        {
            return WorkbookRecoveryRefusalReason.ConfigurationRenameFailed;
        }

        ProtectionGuardOutcome protection = _protectionGuard.QueryTarget(sheet);
        if (protection != ProtectionGuardOutcome.NotProtected)
        {
            return protection == ProtectionGuardOutcome.NoActiveWorkbook
                ? WorkbookRecoveryRefusalReason.NoActiveWorkbook
                : WorkbookRecoveryRefusalReason.TargetProtected;
        }

        string previousName = location.SheetName;
        try
        {
            SetSheetName(sheet, GanttWorkbookContract.ConfigSheetName);
            SetSheetVisibility(sheet, Excel.XlSheetVisibility.xlSheetVeryHidden);

            ConfigWriteOutcome write = _catalogueWriter.Write();
            if (!write.Succeeded)
            {
                RestoreSheetName(sheet, previousName);
                return WorkbookRecoveryRefusalReason.ConfigurationWriteFailed;
            }

            _trace?.Invoke(
                string.Create(
                    CultureInfo.InvariantCulture,
                    $"gantt.recovery renamed configuration sheet={previousName}"));
            return null;
        }
        catch (System.Runtime.InteropServices.COMException exception)
        {
            _trace?.Invoke("gantt.recovery configuration rename refused " + exception.Message);
            RestoreSheetName(sheet, previousName);
            return WorkbookRecoveryRefusalReason.ConfigurationRenameFailed;
        }
    }

    private void RestoreSheetName(Excel.Worksheet sheet, string name)
    {
        try
        {
            SetSheetName(sheet, name);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Best-effort name restore; the typed refusal is the reported outcome.
        }
    }

    private static WorkbookRecoveryRefusalReason? RefusalForConfigState(ConfigSheetState state) =>
        state switch
        {
            ConfigSheetState.Present => null,
            ConfigSheetState.Missing => null,
            ConfigSheetState.Renamed => null,
            ConfigSheetState.NameCollision => WorkbookRecoveryRefusalReason.ConfigurationNameCollision,
            ConfigSheetState.Ambiguous => WorkbookRecoveryRefusalReason.ConfigurationAmbiguous,
            ConfigSheetState.AmbiguousCandidate =>
                WorkbookRecoveryRefusalReason.ConfigurationCandidateAmbiguous,
            ConfigSheetState.Unrecognised => WorkbookRecoveryRefusalReason.ConfigurationUnrecognised,
            _ => null,
        };

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
                && GanttWorkbookContract.IsConfigSheetName(GetSheetName(worksheet)))
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
                && GanttWorkbookContract.IsGanttSheetLabel(GetSheetName(worksheet)))
            {
                labelled.Add(worksheet);
            }
        }

        return labelled;
    }

    private bool TryFindSheetByName(
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
                && string.Equals(GetSheetName(candidate), name, StringComparison.OrdinalIgnoreCase))
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
                if (GanttChartContract.IsDataTableName(GetTableName(candidateTable)))
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

    private Excel.ListObject? GetSingleTable(Excel.Worksheet worksheet)
    {
        Excel.ListObjects objects = worksheet.ListObjects;
        return GetTableCount(objects) == 1 ? GetTableAt(objects, 1) : null;
    }

    private bool HasPlotAnchor(Excel.Worksheet target)
    {
        Excel.Names names = target.Names;
        int count = GetNameCount(names);
        for (int index = 1; index <= count; index++)
        {
            if (GanttWorkbookContract.IsPlotAnchorName(GetNameText(GetNameAt(names, index))))
            {
                return true;
            }
        }

        return false;
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

    internal virtual string GetSheetName(Excel.Worksheet worksheet) => worksheet.Name;

    internal virtual int GetNameCount(Excel.Names names) => names.Count;

    internal virtual Excel.Name GetNameAt(Excel.Names names, int index) => names.Item(index);

    internal virtual string GetNameText(Excel.Name name) => name.Name;

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
