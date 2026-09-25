using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Core;
using Moq;

namespace GanttCreator.Office.ContractTests;

/// <summary>Positive refusal and zero-mutation tests for TypeOptions materialisation.</summary>
public class TypeOptionsMaterialiserTests
{
    [Fact]
    public void Materialise_refuses_without_an_active_workbook_before_reading_catalogue()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NoActiveWorkbook);
        var catalogue = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            null, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NoActiveWorkbook), outcome);
        catalogue.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Materialise_refuses_protected_workbook_before_catalogue_read()
    {
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.SheetProtected);
        var catalogue = new Mock<IConfigCatalogueReader>(MockBehavior.Strict);

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            null, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.TargetProtected), outcome);
        catalogue.Verify(r => r.Read(), Times.Never);
    }

    [Fact]
    public void Materialise_maps_missing_configuration_to_a_typed_refusal()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.ConfigSheetMissing));

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            application.Object, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.ConfigSheetMissing), outcome);
    }

    [Fact]
    public void Materialise_maps_catalogue_hash_mismatch_to_a_typed_refusal()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        _ = application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        var guard = new Mock<IWorksheetProtectionGuard>();
        _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
        var catalogue = new Mock<IConfigCatalogueReader>();
        _ = catalogue.Setup(r => r.Read()).Returns(ConfigReadOutcome.Refused(ConfigReadRefusalReason.CatalogueHashMismatch));

        TypeOptionsMaterialiseOutcome outcome = new ExcelTypeOptionsMaterialiser(
            application.Object, catalogue.Object, guard.Object).Materialise();

        Assert.Equal(TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.CatalogueHashMismatch), outcome);
    }

    /// <summary>
    /// The name-target guard must accept the value Excel echoes back for a name
    /// this adapter wrote, and must still refuse a name pointing elsewhere.
    /// </summary>
    [Theory]
    // Excel strips the quotes from a sheet name that does not need them, so the
    // adapter's own quoted value is read back unquoted. Accepting that form is the
    // regression: refusing it made every Add Row report TypeOptionsUnavailable.
    [InlineData("=_GanttCreatorConfig!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    [InlineData("='_GanttCreatorConfig'!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    [InlineData("=_ganttcreatorconfig!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", true)]
    // A sheet name containing a quote is stored escaped and read back escaped.
    [InlineData("='It''s Config'!$B$2", "='It''s Config'!$B$2", true)]
    // A real mismatch is still refused.
    [InlineData("=_GanttCreatorConfig!$B$3:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("='Gantt Data'!$B$2:$B$17", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("=Sheet1!$B$2", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("no-separator", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    [InlineData("", "='_GanttCreatorConfig'!$B$2:$B$17", false)]
    public void RefersTo_comparison_accepts_excel_normalisation_and_refuses_a_different_target(
        string existing,
        string expected,
        bool expectedMatch)
    {
        Assert.Equal(expectedMatch, ExcelTypeOptionsMaterialiser.RefersToMatches(existing, expected));
    }

    [Fact]
    public void RefersTo_comparison_refuses_a_missing_stored_target()
    {
        Assert.False(ExcelTypeOptionsMaterialiser.RefersToMatches(null, "='_GanttCreatorConfig'!$B$2:$B$17"));
    }

    /// <summary>
    /// The repair path must be the one caller allowed to replace a name left
    /// pointing at the previous catalogue range: the checker reports the name
    /// invalid, and the repairer's only TypeOptions action would otherwise be
    /// that refusal, leaving the finding unrepairable.
    /// </summary>
    [Fact]
    public void Materialise_for_repair_replaces_a_stale_name_target_while_materialise_refuses()
    {
        var graph = new TypeOptionsGraph("=_GanttCreatorConfig!$B$3:$B$17");

        Assert.Equal(
            TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid),
            graph.Build().Materialise());
        Assert.Equal("=_GanttCreatorConfig!$B$3:$B$17", graph.StoredRefersTo);
        graph.VerifyValidationApplied(Times.Never());

        Assert.Equal(TypeOptionsMaterialiseOutcome.Ok(), graph.Build().MaterialiseForRepair());
        Assert.Equal("='_GanttCreatorConfig'!$B$2:$B$17", graph.StoredRefersTo);
        graph.VerifyValidationApplied(Times.Once());
    }

    /// <summary>
    /// <c>ExcelGanttRowInserter</c> calls this on every Add Row, so a name and
    /// validation that are already current must not be re-applied. The stored
    /// target uses Excel's own unquoted echo of the adapter's value.
    /// </summary>
    [Fact]
    public void Ensure_current_skips_rebuilding_a_current_name_and_validation()
    {
        var graph = new TypeOptionsGraph(
            storedNameRefersTo: "=_GanttCreatorConfig!$B$2:$B$17",
            validationFormula: $"={GanttWorkbookContract.TypeOptionsDefinedName}",
            validationType: Excel.XlDVType.xlValidateList);

        Assert.Equal(TypeOptionsMaterialiseOutcome.Ok(), graph.Build().EnsureCurrent());

        Assert.Equal("=_GanttCreatorConfig!$B$2:$B$17", graph.StoredRefersTo);
        graph.VerifyValidationApplied(Times.Never());
    }

    /// <summary>
    /// A current name whose validation points at some other list must still be
    /// materialised: skipping would leave the dropdown unapplied.
    /// </summary>
    [Fact]
    public void Ensure_current_rebuilds_when_the_validation_is_not_a_list_against_the_name()
    {
        var graph = new TypeOptionsGraph(
            storedNameRefersTo: "=_GanttCreatorConfig!$B$2:$B$17",
            validationFormula: "=$A$1:$A$3",
            validationType: Excel.XlDVType.xlValidateList);

        Assert.Equal(TypeOptionsMaterialiseOutcome.Ok(), graph.Build().EnsureCurrent());

        graph.VerifyValidationApplied(Times.Once());
    }

    /// <summary>
    /// An unreadable validation must not be reported as current: the safe
    /// fallback is to re-apply, which is idempotent.
    /// </summary>
    [Fact]
    public void Ensure_current_rebuilds_when_the_validation_cannot_be_read()
    {
        var graph = new TypeOptionsGraph(storedNameRefersTo: "=_GanttCreatorConfig!$B$2:$B$17");
        graph.Validation.SetupGet(v => v.Type).Throws(new TestComException());

        Assert.Equal(TypeOptionsMaterialiseOutcome.Ok(), graph.Build().EnsureCurrent());

        graph.VerifyValidationApplied(Times.Once());
    }

    /// <summary>
    /// A named <see cref="COMException"/> subclass so the graph can simulate an
    /// Excel read failure without constructing a reserved exception type
    /// directly.
    /// </summary>
    private sealed class TestComException : COMException
    {
        public TestComException()
            : base("Simulated Excel read failure.")
        {
        }

        public TestComException(string message)
            : base(message)
        {
        }

        public TestComException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }

    /// <summary>
    /// Currency is only a reason to skip work, never a reason to overwrite a
    /// name this adapter did not write.
    /// </summary>
    [Fact]
    public void Ensure_current_keeps_the_name_target_refusal_for_a_stale_name()
    {
        var graph = new TypeOptionsGraph("=_GanttCreatorConfig!$B$3:$B$17");

        TypeOptionsMaterialiseOutcome outcome = graph.Build().EnsureCurrent();

        Assert.Equal(
            TypeOptionsMaterialiseOutcome.Refused(TypeOptionsRefusalReason.NameTargetInvalid),
            outcome);
        Assert.Equal("=_GanttCreatorConfig!$B$3:$B$17", graph.StoredRefersTo);
    }

    /// <summary>
    /// A minimal live-shaped COM graph: a workbook that already has the config
    /// sheet, the Gantt sheet, both tables, and a stored TypeOptions name, so
    /// the materialiser reaches the name-target decision instead of an early
    /// refusal.
    /// </summary>
    private sealed class TypeOptionsGraph
    {
        private string? _storedRefersTo;
        private readonly Mock<Excel.Name> _name = new();
        private readonly Mock<Excel.Workbook> _workbook = new();

        public TypeOptionsGraph(
            string? storedNameRefersTo,
            string? validationFormula = null,
            Excel.XlDVType validationType = Excel.XlDVType.xlValidateInputOnly)
        {
            _storedRefersTo = storedNameRefersTo;
            _ = _name.SetupGet(n => n.Name).Returns(GanttWorkbookContract.TypeOptionsDefinedName);
            _ = _name.SetupGet(n => n.RefersTo).Returns(() => _storedRefersTo ?? string.Empty);
            // The PIA types RefersTo as object, so the callback takes object.
            _ = _name.SetupSet(n => n.RefersTo = It.IsAny<object>())
                .Callback<object>(value => _storedRefersTo = value as string);

            Validation = new Mock<Excel.Validation>();
            _ = Validation.SetupGet(v => v.Type).Returns((int)validationType);
            _ = Validation.SetupGet(v => v.Formula1).Returns(validationFormula ?? string.Empty);

            var typeRange = new Mock<Excel.Range>();
            _ = typeRange.SetupGet(r => r.Validation).Returns(Validation.Object);

            var names = new Mock<Excel.Names>();
            _ = names.SetupGet(n => n.Count).Returns(1);
            _ = names.Setup(n => n.Item(1)).Returns(_name.Object);
            _ = _workbook.SetupGet(w => w.Names).Returns(names.Object);

            var sheets = new Mock<Excel.Sheets>();
            _ = _workbook.SetupGet(w => w.Sheets).Returns(sheets.Object);
            _ = sheets.SetupGet(s => s.Count).Returns(2);
            Mock<Excel.Worksheet> config = BuildSheet(
                GanttWorkbookContract.ConfigSheetName,
                GanttCatalogues.TypesTableName,
                "DisplayName");
            Mock<Excel.Worksheet> gantt = BuildSheet("Gantt", GanttTableSchema.TableName, "Type");
            _ = sheets.Setup(s => s[1]).Returns(config.Object);
            _ = sheets.Setup(s => s[2]).Returns(gantt.Object);

            // Re-point the Gantt Type column at the range carrying the
            // validation this graph records writes against.
            var typeColumn = new Mock<Excel.ListColumn>();
            _ = typeColumn.SetupGet(c => c.Name).Returns("Type");
            _ = typeColumn.SetupGet(c => c.DataBodyRange).Returns(typeRange.Object);
            var dataColumns = new Mock<Excel.ListColumns>();
            _ = dataColumns.SetupGet(c => c.Count).Returns(1);
            _ = dataColumns.Setup(c => c[1]).Returns(typeColumn.Object);
            var dataTable = new Mock<Excel.ListObject>();
            _ = dataTable.SetupGet(t => t.Name).Returns(GanttTableSchema.TableName);
            _ = dataTable.SetupGet(t => t.ListColumns).Returns(dataColumns.Object);
            var dataObjects = new Mock<Excel.ListObjects>();
            _ = dataObjects.SetupGet(l => l.Count).Returns(1);
            _ = dataObjects.Setup(l => l[1]).Returns(dataTable.Object);
            _ = gantt.SetupGet(w => w.ListObjects).Returns(dataObjects.Object);
        }

        /// <summary>Gets the Type-column validation mock, for write verification.</summary>
        public Mock<Excel.Validation> Validation { get; }

        /// <summary>Gets the TypeOptions name target as currently stored.</summary>
        public string? StoredRefersTo => _storedRefersTo;

        /// <summary>
        /// Verifies how often the adapter re-applied the list validation.
        /// <c>Validation.Add</c> takes five arguments, so it is matched by
        /// arity-free argument matchers rather than a per-argument setup.
        /// </summary>
        /// <param name="times">The expected number of re-applications.</param>
        public void VerifyValidationApplied(Times times) =>
            Validation.Verify(
                v => v.Add(
                    It.IsAny<Excel.XlDVType>(),
                    It.IsAny<Excel.XlDVAlertStyle>(),
                    It.IsAny<object>(),
                    It.IsAny<object>(),
                    It.IsAny<object>()),
                times);

        private static Mock<Excel.Worksheet> BuildSheet(string name, string tableName, string columnName)
        {
            var worksheet = new Mock<Excel.Worksheet>();
            _ = worksheet.SetupGet(w => w.Name).Returns(name);

            var column = new Mock<Excel.ListColumn>();
            _ = column.SetupGet(c => c.Name).Returns(columnName);
            _ = column.SetupGet(c => c.DataBodyRange).Returns(new Mock<Excel.Range>().Object);
            var columns = new Mock<Excel.ListColumns>();
            _ = columns.SetupGet(c => c.Count).Returns(1);
            _ = columns.Setup(c => c[1]).Returns(column.Object);
            var table = new Mock<Excel.ListObject>();
            _ = table.SetupGet(t => t.Name).Returns(tableName);
            _ = table.SetupGet(t => t.ListColumns).Returns(columns.Object);
            var listObjects = new Mock<Excel.ListObjects>();
            _ = listObjects.SetupGet(l => l.Count).Returns(1);
            _ = listObjects.Setup(l => l[1]).Returns(table.Object);
            _ = worksheet.SetupGet(w => w.ListObjects).Returns(listObjects.Object);
            return worksheet;
        }

        /// <summary>Builds the adapter over this graph.</summary>
        /// <returns>The materialiser under test.</returns>
        public ExcelTypeOptionsMaterialiser Build()
        {
            var application = new Mock<Excel.Application>();
            _ = application.SetupGet(a => a.ActiveWorkbook).Returns(_workbook.Object);
            var guard = new Mock<IWorksheetProtectionGuard>();
            _ = guard.Setup(g => g.Query()).Returns(ProtectionGuardOutcome.NotProtected);
            _ = guard.Setup(g => g.QueryTarget(It.IsAny<Excel.Worksheet>()))
                .Returns(ProtectionGuardOutcome.NotProtected);
            var catalogue = new Mock<IConfigCatalogueReader>();
            _ = catalogue.Setup(r => r.Read()).Returns(
                ConfigReadOutcome.Ok(
                    "workbook",
                    new Dictionary<string, string>(),
                    GanttStyleRegistry.Empty));
            return new TestableTypeOptionsMaterialiser(application.Object, catalogue.Object, guard.Object);
        }

        /// <summary>
        /// Substitutes the parameterised <c>Range.Address</c> read, which an
        /// expression tree cannot contain, with the catalogue address the
        /// production code would build.
        /// </summary>
        private sealed class TestableTypeOptionsMaterialiser(
            object? application,
            IConfigCatalogueReader catalogueReader,
            IWorksheetProtectionGuard protectionGuard)
            : ExcelTypeOptionsMaterialiser(application, catalogueReader, protectionGuard)
        {
            internal override string GetRefersTo(string sheetName, Excel.Range range) =>
                $"='{sheetName.Replace("'", "''", StringComparison.Ordinal)}'!$B$2:$B$17";
        }
    }
}
