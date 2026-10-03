#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable Step-0 probe for the Add-Activity row-insert defect.

.DESCRIPTION
    Two questions decide the fix, and neither is answerable from source or from
    Microsoft documentation:

      Q1. What is Shape.Placement immediately after creation, for EACH family the
          renderer actually uses -- AddShape (bars), AddLine (delineators and grid
          lines), AddTextbox (labels), and the msoShapeDiamond milestone?

          ADR-0035 probed AddTextbox ONLY and concluded xlMoveAndSize, then wrote
          no Placement at all. The bars are AddShape and the delineators are
          AddLine, so two of the four families are unmeasured.

      Q2. When a real worksheet row is inserted above a shape, does that shape's
          Top move down by exactly one row height, and does TopLeftCell follow?

    Also measures the REORDER under discussion in the fix plan: insert the real
    worksheet row FIRST, then call ListRows.Add(), and read back the resulting
    body-row and padding-row heights. The plan assumes the fresh row inherits the
    body height from the row above and the padding row keeps its own 6pt.

.NOTES
    Disposable. Per AGENTS.md this is a spike whose output is recorded in the ADR
    and then discarded; it is not a committed test. Run with:

        pwsh ./scripts/probe-rowinsert-anchoring.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121              # xlShiftDown
$xlFormatFromLeftOrAbove = -4142  # xlFormatFromLeftOrAbove

$bodyHeightPt = 18.0    # GanttRowHeightPt default
$paddingHeightPt = 6.0  # ChartPaddingRowHeightPt default
$headerHeightPt = 16.0  # PeriodBandHeightPt default

# 'Build-' rather than 'New-': PSScriptAnalyzer's PSUseShouldProcessForStateChangingFunctions
# flags a New- verb as state-changing and demands ShouldProcess support, which a probe
# helper that only builds a throwaway in-memory sheet has no use for.
function Build-ProbeSheet {
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] $Workbook
    )

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = 'Probe'

    # A table standing in for tblGanttData: header on row 1, three body rows.
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)  # xlSrcRange, xlYes
    $table.Name = 'tblProbe'

    # Rows is a parameterized COM property; .Item(n) is the reliable accessor.
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..4) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }

    # The reserved bottom padding row the chart uses as its bottom margin.
    $sheet.Rows.Item(5).RowHeight = $paddingHeightPt

    return @{ Sheet = $sheet; Table = $table }
}

function Show-ShapeAnchoring {
    param(
        [Parameter(Mandatory)] $Shape,
        [Parameter(Mandatory)] [string] $Label
    )

    $placement = $Shape.Placement
    $placementName = switch ($placement) {
        1 { 'xlMoveAndSize' }
        2 { 'xlMove' }
        3 { 'xlFreeFloating' }
        default { "UNKNOWN($placement)" }
    }

    $topLeft = $Shape.TopLeftCell.Address($false, $false)
    $bottomRight = $Shape.BottomRightCell.Address($false, $false)

    '{0,-22} Placement={1,-14} Top={2,-9} TopLeftCell={3,-6} BottomRightCell={4}' -f
        $Label, $placementName, $Shape.Top, $topLeft, $bottomRight
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "Excel version: $($excel.Version)"
    Write-Host ''

    $workbook = $excel.Workbooks.Add()
    $ctx = Build-ProbeSheet -Excel $excel -Workbook $workbook
    $sheet = $ctx.Sheet
    $table = $ctx.Table

    # ---------------------------------------------------------------- Q1
    Write-Host '=== Q1: Placement immediately after creation, per family ==='

    # Bars are placed on the worksheet row they belong to, matching what
    # ExcelShapeWriter.AddShape does: an absolute point rectangle.
    $shapes = $sheet.Shapes
    $bar = $shapes.AddShape(1, 200.0, 60.0, 90.0, 8.0)            # msoShapeRectangle
    $line = $shapes.AddLine(200.0, 80.0, 290.0, 80.0)
    $label = $shapes.AddTextbox(1, 300.0, 60.0, 70.0, 12.0)         # msoTextOrientationHorizontal
    $diamond = $shapes.AddShape(4, 400.0, 60.0, 8.0, 8.0)          # msoShapeDiamond

    Show-ShapeAnchoring -Shape $bar -Label 'AddShape (bar)'
    Show-ShapeAnchoring -Shape $line -Label 'AddLine (delineator)'
    Show-ShapeAnchoring -Shape $label -Label 'AddTextbox (label)'
    Show-ShapeAnchoring -Shape $diamond -Label 'AddShape (milestone)'
    Write-Host ''

    Write-Host '=== Q2: does a real worksheet row insert move the shape? ==='

    # Snapshot before the insert.
    $before = @{}
    foreach ($pair in @(
            @('bar', $bar), @('line', $line),
            @('label', $label), @('diamond', $diamond))) {
        $before[$pair[0]] = @{
            Top = $pair[1].Top
            TopLeft = $pair[1].TopLeftCell.Address($false, $false)
            Placement = $pair[1].Placement
        }
    }

    # Insert one genuine worksheet row at row 3 -- ABOVE the bar's anchor and
    # inside the table's range, which is the case Excel reportedly refuses.
    $inserted = $false
    try {
        $sheet.Rows.Item(3).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
        $inserted = $true
        Write-Host "Rows[3].Insert inside the ListObject range: SUCCEEDED"
    }
    catch {
        Write-Host "Rows[3].Insert inside the ListObject range: REFUSED -- $($_.Exception.Message)"
    }

    if ($inserted) {
        foreach ($pair in @(
                @('bar', $bar), @('line', $line),
                @('label', $label), @('diamond', $diamond))) {
            $name = $pair[0]
            $shape = $pair[1]
            $delta = $shape.Top - $before[$name].Top
            '{0,-22} Top {1,-9} -> {2,-9} delta={3,-7} TopLeftCell {4,-6} -> {5}' -f
                $name, $before[$name].Top, $shape.Top, $delta,
                $before[$name].TopLeft, $shape.TopLeftCell.Address($false, $false)
        }
    }

    Write-Host ''
    Write-Host '=== Q3: the REORDER -- worksheet row first, then ListRows.Add() ==='

    # Fresh sheet so Q2's insert does not perturb the measurement.
    $sheet2 = $workbook.Worksheets.Add()
    $sheet2.Name = 'Probe2'
    $table2 = $sheet2.ListObjects.Add(1, $sheet2.Range('A1:C4'), $null, 1)
    $table2.Name = 'tblProbe2'
    $sheet2.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..4) { $sheet2.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet2.Rows.Item(5).RowHeight = $paddingHeightPt

    $table2LastRowBefore = $table2.Range.Row + $table2.Range.Rows.Count - 1
    Write-Host "table last row before : $table2LastRowBefore"
    Write-Host "body heights before   : $((2..4 | ForEach-Object { $sheet2.Rows.Item($_).RowHeight }) -join ', ')"
    Write-Host "padding height before : $($sheet2.Rows.Item(5).RowHeight)"

    # THE REORDER: real row first.
    $sheet2.Rows.Item($table2LastRowBefore + 1).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    # ...then let the table claim the fresh row.
    $newRow = $table2.ListRows.Add([Type]::Missing)

    $newRowIndex = $newRow.Index
    Write-Host "ListRows.Add() returned body index: $newRowIndex"
    Write-Host "table last row after  : $($table2.Range.Row + $table2.Range.Rows.Count - 1)"
    Write-Host "body row $newRowIndex height : $($newRow.Range.RowHeight)"
    Write-Host "row below table height      : $($sheet2.Rows.Item($newRowIndex + 2).RowHeight)"
    Write-Host "row below that height       : $($sheet2.Rows.Item($newRowIndex + 3).RowHeight)"

    Write-Host ''
    Write-Host '=== Q4: the CURRENT order -- ListRows.Add() first, row after ==='

    $sheet3 = $workbook.Worksheets.Add()
    $sheet3.Name = 'Probe3'
    $table3 = $sheet3.ListObjects.Add(1, $sheet3.Range('A1:C4'), $null, 1)
    $table3.Name = 'tblProbe3'
    $sheet3.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..4) { $sheet3.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet3.Rows.Item(5).RowHeight = $paddingHeightPt

    # The order the shipped code uses.
    $newRow3 = $table3.ListRows.Add([Type]::Missing)
    $newRow3Index = $newRow3.Index
    Write-Host "ListRows.Add() returned body index: $newRow3Index"
    Write-Host "body row $newRow3Index height (want $bodyHeightPt) : $($newRow3.Range.RowHeight)"
    Write-Host "table last row after  : $($table3.Range.Row + $table3.Range.Rows.Count - 1)"

    $lastRow3 = $table3.Range.Row + $table3.Range.Rows.Count - 1
    $sheet3.Rows.Item($lastRow3 + 1).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    $sheet3.Rows.Item($lastRow3 + 1).RowHeight = $paddingHeightPt
    Write-Host "after the compensating insert:"
    Write-Host "  new body row height   : $($table3.ListRows.Item($newRow3Index).Range.RowHeight)"
    Write-Host "  padding row height    : $($sheet3.Rows.Item($lastRow3 + 1).RowHeight)"

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