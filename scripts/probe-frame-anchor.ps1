#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: can a plot-band shape be ANCHORED so it always stretches?

.DESCRIPTION
    probe-frame-stretch.ps1 measured that xlMoveAndSize resizes a shape only when the
    insertion point is STRICTLY BELOW the shape's TopLeftCell row. An insert at the
    shape's own top row slides the shape down and leaves its height alone, which is
    why a row added at the top of the body leaves the plot unpainted there.

    Excel derives TopLeftCell from the shape's Top, so the only lever the host offers
    is WHERE the shape's top edge sits. The candidate fix: anchor the plot bands into
    the HEADER row (a hair inside it) instead of starting exactly on the first body
    row's boundary, and paint them BEHIND the header so the overlap is invisible. Then
    every body insert is strictly below the anchor and the shape must resize.

    Compares baseline against candidate across inserts at the top, middle and below.

.NOTES
    Disposable spike. Results are printed as they are measured rather than returned,
    because values built from COM proxies do not survive a PowerShell function
    boundary reliably. Run with:

        pwsh ./scripts/probe-frame-anchor.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0
$overlapPt = 0.5

$headerFormat = '{0,-12} {1,-6} {2,-10} {3,-12} {4,-12} {5,-12} {6}'

function Measure-Case {
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] [string] $Case,
        [Parameter(Mandatory)] [double] $ShapeTop,
        [Parameter(Mandatory)] [int] $InsertRow
    )

    $book = $Excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = "s$InsertRow$($Case.Substring(0, 1))"

    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C8'), $null, 1)
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..8) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(9).RowHeight = $paddingHeightPt

    # The plot-background analogue, spanning the body rows.
    $shape = $sheet.Shapes.AddShape(1, 200.0, $ShapeTop, 300.0, 200.0)

    $beforeTop = $shape.Top
    $beforeHeight = $shape.Height
    $beforeBottom = $shape.Top + $shape.Height
    $beforeAnchor = $shape.TopLeftCell.Address($false, $false)

    $sheet.Rows.Item($InsertRow).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)

    $afterAnchor = $shape.TopLeftCell.Address($false, $false)
    $topDelta = [Math]::Round($shape.Top - $beforeTop, 2)
    $heightDelta = [Math]::Round($shape.Height - $beforeHeight, 2)
    $bottomDelta = [Math]::Round(($shape.Top + $shape.Height) - $beforeBottom, 2)

    $outcome = if ([Math]::Abs($heightDelta - $bodyHeightPt) -lt 0.01) { 'stretches OK' }
               elseif ([Math]::Abs($heightDelta) -lt 0.01) { 'SLIDES (bug)' }
               else { "delta=$heightDelta" }

    Write-Host ($headerFormat -f $Case, $InsertRow, $topDelta, $heightDelta, $bottomDelta, "$beforeAnchor->$afterAnchor", $outcome)

    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
}

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"

    $baselineTop = $headerHeightPt
    $candidateTop = $headerHeightPt - $overlapPt

    Write-Host ''
    Write-Host ("baseline Top  = {0}pt (exactly on the header/body boundary)" -f $baselineTop)
    Write-Host ("candidate Top = {0}pt ({1}pt inside the header row)" -f $candidateTop, $overlapPt)
    Write-Host ''

    Write-Host ($headerFormat -f 'Case', 'Row', 'TopDelta', 'HeightDelta', 'BottomDelta', 'Anchor', 'Outcome')

    foreach ($row in @(2, 5, 9)) {
        Measure-Case -Excel $excel -Case 'baseline' -ShapeTop $baselineTop -InsertRow $row
    }
    foreach ($row in @(2, 5, 9)) {
        Measure-Case -Excel $excel -Case 'CANDIDATE' -ShapeTop $candidateTop -InsertRow $row
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