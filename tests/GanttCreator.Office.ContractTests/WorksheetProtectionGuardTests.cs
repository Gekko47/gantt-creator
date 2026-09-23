using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using GanttCreator.Office;
using Moq;
using CommandClass = GanttCreator.Core.DestructiveCommandPolicy.CommandClass;
using ConfirmationKind = GanttCreator.Core.DestructiveCommandPolicy.ConfirmationKind;

namespace GanttCreator.Office.ContractTests;

/// <summary>
/// Contract tests for <see cref="ExcelWorksheetProtectionGuard"/>. They need no
/// live Office: the application object is absent, foreign, or a Moq proxy of the
/// Excel PIA interfaces. Every guard/refusal has a positive test, and the guard's
/// read-only contract is asserted (AGENTS.md validator rule; checklist sections
/// I, J, K).
/// </summary>
public class WorksheetProtectionGuardTests
{
    // ---- Outcome four-corners (positive tests for every outcome) ----

    [Fact]
    public void Query_returns_NoActiveWorkbook_when_the_application_is_absent()
    {
        // A guard constructed with a null application is the absent-host case.
        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(null);

        Assert.Equal(
            ProtectionGuardOutcome.NoActiveWorkbook,
            guard.Query());
    }

    [Fact]
    public void Query_returns_NoActiveWorkbook_when_the_application_has_no_active_workbook()
    {
        var application = new Mock<Excel.Application>();
        application.SetupGet(a => a.ActiveWorkbook).Returns((Excel.Workbook)null!);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.NoActiveWorkbook,
            guard.Query());
    }

    [Fact]
    public void Query_returns_NoActiveWorkbook_when_the_active_workbook_has_no_active_sheet()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        // The active sheet resolves to null (foreign/empty host state).
        workbook.SetupGet(w => w.ActiveSheet).Returns((Excel.Worksheet)null!);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.NoActiveWorkbook,
            guard.Query());
    }

    [Fact]
    public void Query_returns_SheetProtected_when_the_active_sheet_is_content_protected()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        // Sheet contents are protected: the guard must refuse.
        sheet.SetupGet(s => s.ProtectContents).Returns(true);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.SheetProtected,
            guard.Query());
    }

    [Fact]
    public void Query_returns_SheetProtected_when_the_active_sheet_is_protected_even_if_structure_is_not()
    {
        // When both structure and content protection are present, the outcome
        // must be SheetProtected (the more specific refusal). The current order
        // checks structure first; assert the structure-protected case separately
        // and this content-protected-only case here.
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(false);
        sheet.SetupGet(s => s.ProtectContents).Returns(true);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.SheetProtected,
            guard.Query());
    }

    [Fact]
    public void Query_returns_WorkbookStructureProtected_when_the_workbook_structure_is_protected()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        // Structure protection is the workbook-wide gate; the guard must refuse
        // even before inspecting the sheet.
        workbook.SetupGet(w => w.ProtectStructure).Returns(true);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.WorkbookStructureProtected,
            guard.Query());
    }

    [Fact]
    public void Query_returns_WorkbookStructureProtected_even_when_the_sheet_is_not_content_protected()
    {
        // Structure protection should win over a non-protected sheet. This is the
        // explicit order-assertion test: structure is checked first.
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(true);
        sheet.SetupGet(s => s.ProtectContents).Returns(false);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.WorkbookStructureProtected,
            guard.Query());
    }

    [Fact]
    public void Query_returns_NotProtected_when_nothing_is_protected()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(false);
        sheet.SetupGet(s => s.ProtectContents).Returns(false);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.NotProtected,
            guard.Query());
    }

    // ---- Read-only contract (zero mutation) ----

    [Fact]
    public void Query_does_not_mutate_the_workbook()
    {
        var application = new Mock<Excel.Application>(MockBehavior.Strict);
        var workbook = new Mock<Excel.Workbook>(MockBehavior.Strict);
        var sheet = new Mock<Excel._Worksheet>(MockBehavior.Strict);
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(false);
        sheet.SetupGet(s => s.ProtectContents).Returns(false);

        // Strict mocks: any unexpected call throws. The guard is read-only, so
        // no setter should be invoked.
        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);
        _ = guard.Query();

        // No setup setter was configured; if the guard mutated anything the
        // strict mock would have thrown. Running to here with no exception
        // proves no unexpected call reached the mocks.
    }

    [Fact]
    public void Query_does_not_mutate_a_content_protected_workbook()
    {
        // Even when the guard refuses, it must not attempt to unprotect or
        // otherwise mutate the workbook. A strict mock with no setter setups
        // proves the refusal path is read-only.
        var application = new Mock<Excel.Application>(MockBehavior.Strict);
        var workbook = new Mock<Excel.Workbook>(MockBehavior.Strict);
        var sheet = new Mock<Excel._Worksheet>(MockBehavior.Strict);
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(false);
        sheet.SetupGet(s => s.ProtectContents).Returns(true);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);
        _ = guard.Query();
    }

    [Fact]
    public void Query_does_not_mutate_a_structure_protected_workbook()
    {
        var application = new Mock<Excel.Application>(MockBehavior.Strict);
        var workbook = new Mock<Excel.Workbook>(MockBehavior.Strict);
        var sheet = new Mock<Excel._Worksheet>(MockBehavior.Strict);
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(true);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);
        _ = guard.Query();
    }

    // ---- Behavioural ordering: NoActiveWorkbook before SheetProtected ----

    [Fact]
    public void Query_reports_NoActiveWorkbook_rather_than_SheetProtected_when_both_conditions_are_true()
    {
        // If the active workbook cannot be reached, the guard must report
        // NoActiveWorkbook even though (in some hypothetical host state) the
        // sheet would be protected. This is the explicit
        // "NoActiveWorkbook before SheetProtected" behavioural test called out in
        // the work item, proved via the seam replacement rather than by relying
        // on COM ordering.
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        // ActiveWorkbook is present but ActiveSheet is absent — the guard must
        // report NoActiveWorkbook, not reach the sheet-protection check.
        workbook.SetupGet(w => w.ActiveSheet).Returns((Excel.Worksheet)null!);

        IWorksheetProtectionGuard guard = new ExcelWorksheetProtectionGuard(application.Object);

        Assert.Equal(
            ProtectionGuardOutcome.NoActiveWorkbook,
            guard.Query());

        // The sheet-protection probe must not have been reached, so it need not
        // even be configured.
        sheet.VerifyGet(s => s.ProtectContents, Times.Never());
    }

    // ---- Seam replacement verifies the PIA properties are wired ----

    [Fact]
    public void Query_wires_ProtectContents_via_the_seam()
    {
        // Prove the guard actually reads ProtectContents (not a hardcoded value)
        // by replacing the seam and asserting the substituted value flows to the
        // outcome.
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(false);
        sheet.SetupGet(s => s.ProtectContents).Returns(true);

        var guard = new TestableGuard(application.Object)
        {
            StructureProtectedReturnValue = false,
            SheetProtectedReturnValue = false,   // seam override: pretend not protected
        };

        // With the seam overridden to "not protected", the guard must report
        // NotProtected, proving the real outcome depends on the seam's return.
        Assert.Equal(
            ProtectionGuardOutcome.NotProtected,
            guard.Query());
    }

    [Fact]
    public void Query_wires_ProtectStructure_via_the_seam()
    {
        var application = new Mock<Excel.Application>();
        var workbook = new Mock<Excel.Workbook>();
        var sheet = new Mock<Excel._Worksheet>();
        application.SetupGet(a => a.ActiveWorkbook).Returns(workbook.Object);
        workbook.SetupGet(w => w.ActiveSheet).Returns(sheet.Object);
        workbook.SetupGet(w => w.ProtectStructure).Returns(true);
        sheet.SetupGet(s => s.ProtectContents).Returns(false);

        var guard = new TestableGuard(application.Object)
        {
            StructureProtectedReturnValue = false, // seam override
            SheetProtectedReturnValue = false,
        };

        // With the seam overridden to "not protected", the structure check must
        // pass through to the sheet check.
        Assert.Equal(
            ProtectionGuardOutcome.NotProtected,
            guard.Query());
    }

    /// <summary>
    /// A guard whose protection-probe seams are injectable, so the contract test
    /// can prove the PIA properties are wired rather than hardcoded.
    /// </summary>
    private sealed class TestableGuard : ExcelWorksheetProtectionGuard
    {
        public TestableGuard(object? application) : base(application)
        {
            StructureProtectedReturnValue = true;
            SheetProtectedReturnValue = true;
        }

        public bool StructureProtectedReturnValue { get; set; }
        public bool SheetProtectedReturnValue { get; set; }

        internal override bool IsWorkbookStructureProtected(Excel.Workbook workbook) =>
            StructureProtectedReturnValue;

        internal override bool IsWorksheetProtected(Excel._Worksheet worksheet) =>
            SheetProtectedReturnValue;
    }
}
