#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: the R5.4 collapsed-parent lane-anchor defect and its
    fix -- can a row's Hidden flag be written, and does it follow content?

.DESCRIPTION
    Reported defect: "moving up or down a parent of a group when collapsed
    created a lane anchor issue as the parent moves into a hidden row".

    The mechanism, read from source:

      1. A move is a CONTENT rotation (R5.4 D11). `BlockSwap` reorders the
         in-memory cell matrix and `WriteRows` writes the union span back.
         The worksheet ROWS never move.
      2. `Hidden` is a property of the worksheet ROW, not of the content in
         it. So a collapsed group's parent content can land on a row that is
         still Hidden = $true -- it was a hidden child row of the group it
         rotated past.
      3. Refresh measures that row: `ExcelPanelGridMeasurement` reads
         `RowHeight`, which Excel reports as 0 for a hidden row.
      4. `PanelCellGrid` accepts zero deliberately (ADR-0034: a collapse must
         not fail measurement).
      5. `LaneRowAnchorResolver` then refuses with `DegenerateRowHeight`,
         and `SceneBuilder` surfaces `UnresolvableLaneAnchor`. The whole
         Refresh refuses.

    The fix must rotate the Hidden flags WITH the content, so a parent stays
    visible and a collapsed child stays hidden wherever the group lands.
    Unhiding outright would expand the user's collapsed group, which
    `IOutlineGroupPort` deliberately never does.

      Q1. Does writing cell CONTENT into a hidden row change its Hidden flag
          or its measured RowHeight? (expected: no -- Hidden is a row
          property, so the defect is real)
      Q2. Is RowHeight genuinely 0 for a hidden row -- the measurement the
          anchor resolver sees and refuses on?
      Q3. Can `Range.Hidden` be WRITTEN, and through WHICH accessor? This is
          the question that must not be assumed: ADR-0034 recorded a
          `Range[...]`-sourced range REFUSING an `OutlineLevel` write with
          COMException 0x800A03EC while the identically-shaped `Rows[...]`
          range accepted it. Hidden may behave the same way.
      Q4. After rotating the Hidden flags with the content, does the
          presentation match the hierarchy at the new positions -- parent
          visible, collapsed children hidden, heights measured?

    Disposable by design: it drives a throwaway workbook through COM and
    prints measurements. It is not a test, is not run by verify.ps1, and its
    findings are recorded as evidence in the R5.4 work item.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$headerHeightPt = 16.0
$bodyHeightPt = 18.0
$paddingHeightPt = 6.0

function Show-Height {
    param($Row, [string] $Label)

    Write-Host ("  {0,-22} Hidden={1,-5} RowHeight={2}" -f $Label, $Row.Hidden, $Row.RowHeight)
}

function Invoke-ProbeWrite {
    param([string] $What, [scriptblock] $Body)

    try {
        & $Body
        Write-Host ("  {0,-46} ACCEPTED" -f $What)
    }
    catch [System.Runtime.InteropServices.COMException] {
        Write-Host ("  {0,-46} REFUSED  COM 0x{1:X8}" -f $What, $_.Exception.HResult)
    }
    catch {
        Write-Host ("  {0,-46} ERROR    {1}" -f $What, $_.Exception.GetType().Name)
    }
}

$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host

    # header row 1 (16pt) + 5 body rows 2..6 (18pt) + padding row 7 (6pt).
    #
    #   row 2  P1   parent, VISIBLE
    #   row 3  C1   child of P1, HIDDEN (P1 collapsed)
    #   row 4  C2   child of P1, HIDDEN (P1 collapsed)
    #   row 5  P2   parent, VISIBLE
    #   row 6  C3   child of P2, HIDDEN (P2 collapsed)
    $sheet = $excel.Workbooks.Add().Worksheets.Item(1)
    $sheet.Name = 'Hidden'
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C6'), $null, 1)
    $table.Name = 'tblHidden'

    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..6) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(7).RowHeight = $paddingHeightPt

    foreach ($r in 3, 4, 6) {
        $sheet.Rows.Item($r).OutlineLevel = 2
        $sheet.Rows.Item($r).Hidden = $true
    }

    Write-Host '=== BEFORE: the collapsed sheet as laid out ==='
    foreach ($r in 2..7) { Show-Height ($sheet.Rows.Item($r)) ("row $r") }
    Write-Host

    Write-Host '=== Q1: does writing CONTENT into a hidden row change it? ==='
    # Rotate P1 group (rows 2,3,4) past P2 group (rows 5,6): union is rows 2..6
    # and content becomes P2,C3,P1,C1,C2. Exactly the rotation ApplyMove
    # performs -- cell text only, rows untouched.
    $content = @(
        @('P2', '', ''),
        @('C3', 'P2', ''),
        @('P1', '', ''),
        @('C1', 'P1', ''),
        @('C2', 'P1', '')
    )
    for ($i = 0; $i -lt $content.Count; $i++) {
        $sheet.Range(("A{0}:C{0}" -f (2 + $i))).Value2 = $content[$i]
    }
    Write-Host '  content rotated to P2,C3,P1,C1,C2 -- read the ROWS back:'
    foreach ($r in 2..7) { Show-Height ($sheet.Rows.Item($r)) ("row $r") }
    Write-Host

    Write-Host '=== Q1 VERDICT: is a PARENT now on a hidden (zero-height) row? ==='
    foreach ($pair in @(@('P2', 2), @('P1', 4))) {
        $row = $sheet.Rows.Item($pair[1])
        $cellA = $sheet.Range(("A{0}" -f $pair[1])).Value2
        $verdict = if ([double]$row.RowHeight -le 0) { '*** DEGENERATE: anchor resolver REFUSES here ***' } else { 'measurable, anchor resolves' }
        Write-Host ("  {0} content on row {1} (A{1}={2}) Height={3}  ->  {4}" -f $pair[0], $pair[1], $cellA, $row.RowHeight, $verdict)
    }
    Write-Host

    Write-Host '=== Q2: is RowHeight genuinely 0 for a hidden row? ==='
    Show-Height ($sheet.Rows.Item(3)) 'hidden row 3'
    Show-Height ($sheet.Rows.Item(2)) 'visible row 2'
    Write-Host
    Write-Host '=== Q3: can Range.Hidden be WRITTEN, and through which accessor? ==='
    Write-Host '  (ADR-0034 recorded Range[...] REFUSING OutlineLevel while Rows[...] accepted it)'
    Write-Host

    $probeSheet = $excel.Workbooks.Add().Worksheets.Item(1)
    $probeSheet.Name = 'Accessors'
    $probeSheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..7) { $probeSheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $probeTable = $probeSheet.ListObjects.Add(1, $probeSheet.Range('A1:C6'), $null, 1)

    function Read-Row3 {
        Write-Host ("     read back: Hidden={0} RowHeight={1}" -f $probeSheet.Rows.Item(3).Hidden, $probeSheet.Rows.Item(3).RowHeight)
    }

    Invoke-ProbeWrite 'Rows.Item(3).Hidden = $true' { $probeSheet.Rows.Item(3).Hidden = $true }
    Read-Row3
    Invoke-ProbeWrite 'Rows.Item(3).Hidden = $false' { $probeSheet.Rows.Item(3).Hidden = $false }
    Read-Row3
    Write-Host

    Invoke-ProbeWrite "Range('A3:C3').Hidden = `$true" { $probeSheet.Range('A3:C3').Hidden = $true }
    Read-Row3
    Invoke-ProbeWrite "Range('A3:C3').Hidden = `$false" { $probeSheet.Range('A3:C3').Hidden = $false }
    Read-Row3
    Write-Host

    Invoke-ProbeWrite "Range('A3:C3').EntireRow.Hidden = `$true" { $probeSheet.Range('A3:C3').EntireRow.Hidden = $true }
    Read-Row3
    Invoke-ProbeWrite "Range('A3:C3').EntireRow.Hidden = `$false" { $probeSheet.Range('A3:C3').EntireRow.Hidden = $false }
    Read-Row3
    Write-Host

    Invoke-ProbeWrite 'DataBodyRange.Rows.Item(2).Hidden = $true' { $probeTable.DataBodyRange.Rows.Item(2).Hidden = $true }
    Read-Row3
    Invoke-ProbeWrite 'DataBodyRange.Rows.Item(2).Hidden = $false' { $probeTable.DataBodyRange.Rows.Item(2).Hidden = $false }
    Read-Row3
    Write-Host

    Write-Host '  -- does hiding a row inside a ListObject stay hidden? --'
    Invoke-ProbeWrite 'Rows.Item(4).Hidden = $true (in-table)' { $probeSheet.Rows.Item(4).Hidden = $true }
    Write-Host ("     read back: Hidden={0} RowHeight={1}" -f $probeSheet.Rows.Item(4).Hidden, $probeSheet.Rows.Item(4).RowHeight)
    Invoke-ProbeWrite 'Rows.Item(4).Hidden = $false (restore)' { $probeSheet.Rows.Item(4).Hidden = $false }
    Write-Host

    Write-Host '=== Q4: rotating Hidden WITH the content restores the layout ==='
    # P2,C3,P1,C1,C2 means row 2 visible, row 3 hidden, row 4 visible, 5,6 hidden.
    $rotated = @{ 2 = $false; 3 = $true; 4 = $false; 5 = $true; 6 = $true }
    foreach ($r in 2..6) {
        Invoke-ProbeWrite ("Rows.Item($r).Hidden = `$$($rotated[$r])") { $sheet.Rows.Item($r).Hidden = $rotated[$r] }
    }
    Write-Host '  rows after the Hidden flags were rotated with the content:'
    foreach ($r in 2..6) { Show-Height ($sheet.Rows.Item($r)) ("row $r") }
    Write-Host

    Write-Host '=== VERDICT SUMMARY ==='
    Write-Host '  Q1: a parent whose content lands on a formerly-hidden row measures'
    Write-Host '      RowHeight = 0 -> LaneRowAnchorResolver refuses (DegenerateRowHeight)'
    Write-Host '      -> SceneBuilder.UnresolvableLaneAnchor -> Refresh refuses.'
    Write-Host '  Q3: whichever accessor was ACCEPTED above is the write seam the'
    Write-Host '      rotation must use; a REFUSED accessor must not be relied on.'
    Write-Host '  Q4: rotating Hidden with the content is the fix -- parent visible,'
    Write-Host '      collapsed children hidden, every lane-owning row measurable.'

    foreach ($wb in @($excel.Workbooks)) { $wb.Close($false) }
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    Write-Host
    Write-Host 'probe complete'
}
