#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: why the plot's BOTTOM fails to stretch, and what fixes it.

.DESCRIPTION
    Live report 2026-10-03 (after ADR-0037 fixed the top): the plot's bottom still does
    not stretch when a row is appended.

    Root cause found by reading the measurement code rather than the probe:
    ExcelSceneBuildRequestFactory builds the plot from `measuredGrid.TotalRowHeightPt`,
    and ExcelPanelGridMeasurement populates RowHeightsPt from `table.DataBodyRange` ONLY.
    The reserved padding row is measured separately, as the chart's bottom MARGIN. So the
    plot's bottom edge is the bottom of the LAST BODY ROW -- and because Excel resolves
    an edge lying exactly on a row boundary to the row BELOW, the band's BottomRightCell
    is the PADDING row, which is exactly the row the append branch inserts into. An
    insert at the shape's own bottom-anchor row slides it, the mirror of the top case.

    This measures three candidate fixes for the append insert at that row:
      A. current    -- bottom on the last body row's bottom (0 extension)
      B. +1pt       -- the owner's suggestion
      C. +6.5pt     -- through the padding row and 0.5pt into the row below it

.NOTES
    Disposable spike. Run with:

        pwsh ./scripts/probe-frame-bottom-anchor.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0
$bodyCount = 7
$topLiftPt = 0.5

$headerFormat = '{0,-12} {1,-20} {2,-9} {3,-11} {4,-12} {5,-24}'

function Measure-Case {
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] [string] $Arm,
        [Parameter(Mandatory)] [double] $BottomExtendPt,
        [Parameter(Mandatory)] [int] $InsertRow,
        [Parameter(Mandatory)] [string] $InsertLabel
    )

    $book = $Excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = "s$($InsertRow)$($Arm.Substring(0, 1))"

    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C8'), $null, 1)
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..8) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(9).RowHeight = $paddingHeightPt

    # Header/body boundary lifted by the ADR-0037 top overlap; bottom is the LAST BODY
    # ROW's bottom plus the candidate extension -- which is what the real plot uses.
    $plotTop = $headerHeightPt - $topLiftPt
    $plotBottom = $headerHeightPt + ($bodyCount * $bodyHeightPt) + $BottomExtendPt

    $shape = $sheet.Shapes.AddShape(1, 200.0, $plotTop, 300.0, $plotBottom - $plotTop)

    $beforeAnchor = "$($shape.TopLeftCell.Address($false, $false))/$($shape.BottomRightCell.Address($false, $false))"
    $beforeHeight = $shape.Height

    $sheet.Rows.Item($InsertRow).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)

    $afterAnchor = "$($shape.TopLeftCell.Address($false, $false))/$($shape.BottomRightCell.Address($false, $false))"
    $topDelta = [Math]::Round($shape.Top - $plotTop, 2)
    $heightDelta = [Math]::Round($shape.Height - $beforeHeight, 2)

    $outcome = if ([Math]::Abs($heightDelta - $bodyHeightPt) -lt 0.01) { 'STRETCHES' }
               elseif ([Math]::Abs($heightDelta) -lt 0.01) { 'SLIDES (bug)' }
               else { "delta=$heightDelta" }

    Write-Host ($headerFormat -f $Arm, $InsertLabel, $topDelta, $heightDelta, $outcome, "$beforeAnchor -> $afterAnchor")

    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host ''
    Write-Host "Padding row 9 spans $($headerHeightPt + ($bodyCount * $bodyHeightPt))pt .. $(($headerHeightPt + ($bodyCount * $bodyHeightPt)) + $paddingHeightPt)pt."
    Write-Host 'The append branch inserts at row 9. Row 10 is the row below the chart margin.'
    Write-Host ''
    Write-Host ($headerFormat -f 'Arm', 'Insert', 'TopDelta', 'HeightDelta', 'Outcome', 'Anchor TL/BR')

    foreach ($arm in @(
            @{ Name = 'current(0)'; Ext = 0.0 },
            @{ Name = 'plus1pt'; Ext = 1.0 },
            @{ Name = 'plus6.5pt'; Ext = 6.5 })) {
        Measure-Case -Excel $excel -Arm $arm.Name -BottomExtendPt $arm.Ext -InsertRow 9 -InsertLabel 'row 9 (APPEND)'
    }

    Write-Host ''
    Write-Host 'Sanity: does each arm still handle a TOP insert?'
    foreach ($arm in @(
            @{ Name = 'current(0)'; Ext = 0.0 },
            @{ Name = 'plus1pt'; Ext = 1.0 },
            @{ Name = 'plus6.5pt'; Ext = 6.5 })) {
        Measure-Case -Excel $excel -Arm $arm.Name -BottomExtendPt $arm.Ext -InsertRow 2 -InsertLabel 'row 2 (top)'
    }

    Write-Host ''
    Write-Host 'PROBE COMPLETE'
}
finally {
    if ($excel) { $excel.Quit() }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}