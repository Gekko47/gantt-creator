#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: does a chart-frame shape STRETCH when a row is inserted?

.DESCRIPTION
    Live report 2026-10-03: after Add Activity the lane bars move correctly, but the
    plot's vertical structure (the alternate period bands / plot background) does not
    stretch when a row is added at the TOP or the BOTTOM, so the top of the plot area
    is left unpainted.

    The suspicion is that this is NOT a bug in the insert -- it is what
    Shape.Placement = xlMoveAndSize actually does. Excel resizes a shape only when the
    insertion point falls BETWEEN its TopLeftCell and BottomRightCell rows. An insert
    at or above the top, or at or below the bottom, MOVES the shape and leaves its
    height alone. That would make the plot frame stretch only for a MIDDLE insert,
    which matches the report.

    This measures each case against a table-shaped fixture:
      header row 1, body rows 2..8, reserved padding row 9,
      and one tall rectangle spanning rows 2..8 -- the plot-background analogue.

.NOTES
    Disposable spike. Its output is recorded in the ADR and the script is not a
    committed test. Run with:

        pwsh ./scripts/probe-frame-stretch.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0

# (label, insertRow) -- header=1, body=2..8, padding=9
$scenarios = @(
    @{ Label = 'TOP    (row 2 = first body row = shape top)'; Row = 2 },
    @{ Label = 'UPPER  (row 3)'; Row = 3 },
    @{ Label = 'MIDDLE (row 5)'; Row = 5 },
    @{ Label = 'LOWER  (row 7)'; Row = 7 },
    @{ Label = 'BOTTOM (row 8 = last body row = shape bottom)'; Row = 8 },
    @{ Label = 'BELOW  (row 9 = padding, under the shape)'; Row = 9 }
)

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host ''

    $workbook = $excel.Workbooks.Add()

    foreach ($scenario in $scenarios) {
        $sheet = $workbook.Worksheets.Add()
        $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C8'), $null, 1)
        $table.Name = "tbl$($scenario.Row)"
        $sheet.Rows.Item(1).RowHeight = $headerHeightPt
        foreach ($r in 2..8) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
        $sheet.Rows.Item(9).RowHeight = $paddingHeightPt

        # The plot-background analogue: one tall rectangle over the whole body span,
        # created exactly as ExcelShapeWriter.AddShape creates a band.
        $shape = $sheet.Shapes.AddShape(1, 200.0, 40.0, 300.0, 200.0)

        $beforeTop = $shape.Top
        $beforeHeight = $shape.Height
        $beforeBottom = $shape.Top + $shape.Height
        $beforeTopLeft = $shape.TopLeftCell.Address($false, $false)
        $beforeBottomRight = $shape.BottomRightCell.Address($false, $false)

        $sheet.Rows.Item($scenario.Row).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)

        $afterTop = $shape.Top
        $afterHeight = $shape.Height
        $afterBottom = $shape.Top + $shape.Height

        $stretched = [Math]::Abs(($afterHeight - $beforeHeight) - $bodyHeightPt) -lt 0.01
        $verdict = if ($stretched) { 'STRETCHED +18' }
                   elseif ([Math]::Abs($afterHeight - $beforeHeight) -lt 0.01) { 'moved only (height unchanged)' }
                   else { "height delta $([Math]::Round($afterHeight - $beforeHeight, 2))" }

        Write-Host "--- insert at $($scenario.Label) ---"
        Write-Host ("  Top      {0,-8} -> {1,-8} delta={2}" -f `
            $beforeTop, $afterTop, [Math]::Round($afterTop - $beforeTop, 2))
        Write-Host ("  Height   {0,-8} -> {1,-8} delta={2}" -f `
            $beforeHeight, $afterHeight, [Math]::Round($afterHeight - $beforeHeight, 2))
        Write-Host ("  Bottom   {0,-8} -> {1,-8} delta={2}" -f `
            $beforeBottom, $afterBottom, [Math]::Round($afterBottom - $beforeBottom, 2))
        Write-Host ("  Anchors  TopLeftCell {0} -> {1}   BottomRightCell {2} -> {3}" -f `
            $beforeTopLeft, $shape.TopLeftCell.Address($false, $false), `
            $beforeBottomRight, $shape.BottomRightCell.Address($false, $false))
        Write-Host ("  VERDICT  {0}" -f $verdict)
        Write-Host ''

        $sheet.Delete()
    }

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