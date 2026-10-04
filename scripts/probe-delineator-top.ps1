#Requires -Version 7
<#
.SYNOPSIS
    Disposable probe: does the DELINEATOR line lock its top, or slide?

.DESCRIPTION
    ADR-0037 lifted the bands and the vertical grid lines up into the header row so
    a row added at the top of the body STRETCHES them. The delineator was left on
    the raw plot top, so it is the one plot-spanning shape still anchored exactly
    on the header/body boundary -- the condition ADR-0037 measured to SLIDE.

    This measures the current geometry and the lifted candidate on the same sheet,
    with one row inserted at the top of the body, and prints TopDelta/HeightDelta
    for each. No assertions; the numbers are the finding.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$headerHeightPt = 16.0
$bodyHeightPt = 18.0
$anchorHeightPt = 0.25
$padHeightPt = 5.75
$overlapPt = 0.5          # PlotBandHeaderOverlapPt
$extensionPt = 0.75      # anchor + MajorBoundaryPt/2

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host ''

    # --- CASE 1: the CURRENT geometry, top exactly on the header/body boundary ---
    $book1 = $excel.Workbooks.Add()
    $sheet1 = $book1.Worksheets.Item(1)
    $null = $sheet1.ListObjects.Add(1, $sheet1.Range('A1:C8'), $null, 1)
    $sheet1.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..8) { $sheet1.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet1.Rows.Item(9).RowHeight = $anchorHeightPt
    $sheet1.Rows.Item(10).RowHeight = $padHeightPt

    $plotTop = $headerHeightPt
    $plotBottom = $headerHeightPt + (7 * $bodyHeightPt)
    $current = $sheet1.Shapes.AddLine(200.0, $plotTop, 200.0, $plotBottom + $extensionPt)

    $beforeTop = $current.Top
    $beforeH = $current.Height
    $beforeAnchor = $current.TopLeftCell.Address($false, $false)
    $sheet1.Rows.Item(2).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    Write-Host '=== CURRENT: top exactly on the header/body boundary ==='
    Write-Host ("  anchor {0} -> {1}" -f $beforeAnchor, $current.TopLeftCell.Address($false, $false))
    Write-Host ("  TopDelta    {0}" -f [Math]::Round($current.Top - $beforeTop, 3))
    Write-Host ("  HeightDelta {0}" -f [Math]::Round($current.Height - $beforeH, 3))
    $currentVerdict = if ([Math]::Abs($current.Height - $beforeH) -lt 0.01) { 'SLIDES (bug)' } else { 'STRETCHES' }
    Write-Host ("  verdict     {0}" -f $currentVerdict)
    $book1.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book1)

    # --- CASE 2: the candidate, top lifted into the header row ---
    $book2 = $excel.Workbooks.Add()
    $sheet2 = $book2.Worksheets.Item(1)
    $null = $sheet2.ListObjects.Add(1, $sheet2.Range('A1:C8'), $null, 1)
    $sheet2.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..8) { $sheet2.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet2.Rows.Item(9).RowHeight = $anchorHeightPt
    $sheet2.Rows.Item(10).RowHeight = $padHeightPt

    $candidate = $sheet2.Shapes.AddLine(200.0, $plotTop - $overlapPt, 200.0, $plotBottom + $extensionPt)
    $beforeTop2 = $candidate.Top
    $beforeH2 = $candidate.Height
    $beforeAnchor2 = $candidate.TopLeftCell.Address($false, $false)
    $sheet2.Rows.Item(2).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    Write-Host ''
    Write-Host ("=== CANDIDATE: top lifted {0}pt into the header row ===" -f $overlapPt)
    Write-Host ("  anchor {0} -> {1}" -f $beforeAnchor2, $candidate.TopLeftCell.Address($false, $false))
    Write-Host ("  TopDelta    {0}" -f [Math]::Round($candidate.Top - $beforeTop2, 3))
    Write-Host ("  HeightDelta {0}" -f [Math]::Round($candidate.Height - $beforeH2, 3))
    $candidateVerdict = if ([Math]::Abs($candidate.Height - $beforeH2) -lt 0.01) { 'SLIDES (bug)' } else { 'STRETCHES' }
    Write-Host ("  verdict     {0}" -f $candidateVerdict)
    $book2.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book2)

    Write-Host ''
    Write-Host 'PROBE COMPLETE'
}
finally {
    if ($excel) { $excel.Quit() }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}