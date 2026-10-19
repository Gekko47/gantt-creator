#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: does a TRUE worksheet-row delete (Rows.Delete)
    shift move-anchored Gantt shapes, and what happens to the shape
    anchored in the deleted row?

.DESCRIPTION
    probe-delete-row.ps1 measured `ListRows.Delete(index)`: the
    ListObject shrinks by one row but no worksheet row shifts (every
    shape kept its Top and TopLeftCell, and the row below the table
    kept the 18pt body height instead of the 6pt padding height).
    The owner's R5.4 direction (2026-10-12) asks for the removed
    row's chart elements to follow the delete without a Refresh,
    which requires a genuine worksheet-row delete. This probe measures
    that operation on a live host:

      Q1. Is `Rows.Delete` inside a ListObject range allowed, or
          refused (ADR-0036 recorded Excel refusing a worksheet-row
          INSERT inside a ListObject range)?
      Q2. Does a shape anchored BELOW the deleted row move up exactly
          one body row height (the insert-probe symmetry recorded on
          ExcelGanttRowInserter: AddShape reports Placement =
          xlMoveAndSize and shapes followed an insert)?
      Q3. Does the shape anchored IN the deleted row survive (Excel
          never deletes shapes with their host row) and where does it
          land relative to the survivor that moves up beneath it?
      Q4. Does the ListObject shrink to the correct range, and does
          the reserved 6pt padding row stay directly below the table?
      Q5. What Placement does AddShape report here (the anchoring
          mode any shift depends on)?

    Disposable by design: it drives a throwaway workbook through COM
    and prints measurements. It is not a test, is not run by
    verify.ps1, and its findings are recorded as evidence in the
    R5.4 work item.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0

# XlPlacement values (PIA enum, printed raw because pure COM has no
# assembly to load the enum from): 1 = xlMove, 2 = xlMoveAndSize,
# 3 = xlFreeFloating.
function Show-Placement {
    param($Shape, [string] $Name)

    $placement = $Shape.Placement
    $label = switch ([int]$placement) {
        1 { 'xlMove' }
        2 { 'xlMoveAndSize' }
        3 { 'xlFreeFloating' }
        default { "unknown($placement)" }
    }
    Write-Host ("  {0,-8} Placement={1} ({2})" -f $Name, [int]$placement, $label)
}

function Show-Shape {
    param($Shape, [string] $Name, [hashtable] $Before, [string] $Label)

    Write-Host ("  {0,-8} {1,-6} Top {2,-7} -> {3,-7} delta={4,-7} TopLeftCell={5}" -f `
        $Name, $Label, $Before[$Name], $Shape.Top, ($Shape.Top - $Before[$Name]), $Shape.TopLeftCell.Address($false, $false))
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host

    # header row 1 (16pt) + 3 body rows (2,3,4 @ 18pt) + padding row 5 (6pt).
    $sheet = $excel.Workbooks.Add().Worksheets.Item(1)
    $sheet.Name = 'Shift'
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $table.Name = 'tblShift'
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..4) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(5).RowHeight = $paddingHeightPt

    # One bar per body row, in the plot anchor column (D, the first
    # column past the table), the way the renderer places bars.
    $shapes = $sheet.Shapes
    $null = $shapes.AddShape(1, 200.0, 20.0, 80.0, 8.0).Name = 'barRow2'
    $null = $shapes.AddShape(1, 200.0, 38.0, 80.0, 8.0).Name = 'barRow3'
    $null = $shapes.AddShape(1, 200.0, 56.0, 80.0, 8.0).Name = 'barRow4'

    # Content markers so the probe can see the data rows shift.
    $sheet.Range('A2').Value2 = 'row2'
    $sheet.Range('A3').Value2 = 'row3'
    $sheet.Range('A4').Value2 = 'row4'

    $before = @{}
    foreach ($name in @('barRow2', 'barRow3', 'barRow4')) {
        $shape = $shapes.Item($name)
        $before[$name] = $shape.Top
        Show-Placement $shape $name
        Show-Shape $shape $name $before 'before'
    }

    $last = $table.Range.Row + $table.Range.Rows.Count - 1
    Write-Host ("  before: table={0}..{1} ListRows={2} rowBelowHeight={3}pt" -f `
        $table.Range.Row, $last, $table.ListRows.Count, $sheet.Rows.Item($last + 1).RowHeight)

    Write-Host ''
    Write-Host '=== Q1-Q4: true worksheet-row delete of body row 2 (worksheet row 3) ==='
    try {
        $sheet.Rows.Item(3).Delete()
        Write-Host '  Rows.Item(3).Delete(): accepted'
    }
    catch {
        Write-Host ("  Rows.Item(3).Delete(): REFUSED -- {0}" -f $_.Exception.Message)
        Write-Host '  PROBE COMPLETE (delete refused; no further measurement possible)'
        return
    }

    $last = $table.Range.Row + $table.Range.Rows.Count - 1
    Write-Host ("  after : table={0}..{1} ListRows={2} range={3} rowBelowHeight={4}pt rowBelowIsPadding={5}" -f `
        $table.Range.Row, $last, $table.ListRows.Count, `
        $table.Range.Address($false, $false), $sheet.Rows.Item($last + 1).RowHeight, `
        ($sheet.Rows.Item($last + 1).RowHeight -eq $paddingHeightPt))

    # The data rows themselves: did the content shift up?
    Write-Host ("  content: A2={0} A3={1} A4={2}" -f `
        $sheet.Range('A2').Value2, $sheet.Range('A3').Value2, $sheet.Range('A4').Value2)

    # Every shape the collection still holds, by enumeration: a
    # name lookup cannot tell a deleted shape from a renamed one.
    Write-Host ("  shapes remaining: Count={0}" -f $shapes.Count)
    foreach ($shape in $shapes) {
        Write-Host ("    {0,-8} Top={1,-7} TopLeftCell={2} Placement={3}" -f `
            $shape.Name, $shape.Top, $shape.TopLeftCell.Address($false, $false), [int]$shape.Placement)
    }

    foreach ($name in @('barRow2', 'barRow3', 'barRow4')) {
        try {
            $shape = $shapes.Item($name)
            Show-Shape $shape $name $before 'after '
        }
        catch {
            Write-Host ("  {0,-8} after  GONE -- the host removed it with the deleted row" -f $name)
        }
    }

    Write-Host ''
    Write-Host 'PROBE COMPLETE'
}
finally {
    if ($workbook) { $workbook.Close($false) }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($workbook, $excel)) {
        if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
