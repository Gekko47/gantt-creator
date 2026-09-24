namespace GanttCreator.Office;

/// <summary>
/// Why a worksheet-protection guard query returned the outcome it did.
/// </summary>
/// <remarks>
/// Every outcome is typed and expected — routine, not an exception
/// (docs/02-ARCHITECTURE.md "Error handling"). Unexpected COM failures are not
/// refusals and propagate to the command boundary.
/// </remarks>
public enum ProtectionGuardOutcome
{
    /// <summary>
    /// The workbook's active worksheet is present and not protected; a mutating
    /// command may proceed to its confirmation (if any) and then to mutation.
    /// </summary>
    NotProtected = 0,

    /// <summary>
    /// The active worksheet's contents are protected
    /// (<c>Worksheet.ProtectContents == true</c>); the guard refuses before
    /// any mutation, matching the R2.2 <c>TargetProtected</c> pattern extended
    /// by ADR-0008 D4.
    /// </summary>
    SheetProtected = 1,

    /// <summary>
    /// The active workbook's structure is protected
    /// (<c>Workbook.ProtectStructure == true</c>); the guard refuses because
    /// structure protection blocks sheet-level structural changes the mutating
    /// command may need.
    /// </summary>
    WorkbookStructureProtected = 2,

    /// <summary>
    /// The Excel application object or the active workbook is absent; the guard
    /// refuses with no mutation, matching the existing
    /// <c>NoActiveWorkbook</c> refusal family.
    /// </summary>
    NoActiveWorkbook = 3,
}
