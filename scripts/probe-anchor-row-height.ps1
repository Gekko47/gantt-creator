#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe (ADR-0038 Step 0): can Excel hold a 0.5pt row, and does the
    anchor-row layout stretch?

.DESCRIPTION
    Two unknowns before any layout code moves.

    Q1 -- THE ROW-HEIGHT CLAMP. The design calls for a 0.5pt anchor row. Excel clamps
    Range.RowHeight, and the floor is not documented. This SETS candidate heights and
    READS THEM BACK, because a written height is not the stored height.

    Q2 -- DOES THE LAYOUT STRETCH? Built exactly as agreed:
        last activity row 18pt  <- plot bottom
        anchor row        ~0.5pt <- insert target (one past the body)
        frame line (1pt, centred on the anchor/pad boundary)
        pad row           5pt    <- bands extend 0.5pt in, covered by the line
    The band's bottom therefore sits 1.0pt below the plot bottom. The question is
    whether inserting at the ANCHOR row stretches it (the bottom-anchor case), which
    is the opposite behaviour to the top-anchor case already measured.

.NOTES
    Disposable spike. Run with:

        pwsh ./scripts/probe-anchor-row-height.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$bodyHeightPt = 18.0
$padHeightPt = 5.0
$frameWidthPt = 1.0
$overhangPt = 0.5
$headerHeightPt = 16.0

$headerFormat = '{0,-14} {1,-10} {2,-10} {3,-12} {4,-12}'

function Measure-Case {
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] [double] $AnchorHeightPt,
        [Parameter(Mandatory)] [int] $InsertRow,
        [Parameter(Mandatory)] [string] $InsertLabel
    )

    $book = $Excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    $sheet.Name = "s$($InsertRow)$([int]($AnchorHeightPt * 10))"

    # rows: 1 header, 2..8 body, 9 anchor, 10 pad
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C8'), $null, 1)
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..8) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(9).RowHeight = $AnchorHeightPt      # the anchor row
    $sheet.Rows.Item(10).RowHeight = $padHeightPt

    $anchorActual = $sheet.Rows.Item(9).RowHeight

    # Plot bottom = bottom of the last body row = top of the anchor row.
    $plotBottom = $headerHeightPt + (7 * $bodyHeightPt)
    # The frame line is 1pt centred on the anchor/pad boundary, and the band must
    # stop exactly where that line stops covering: half a line below it.
    $bandBottom = $plotBottom + $anchorActual + ($frameWidthPt / 2)

    $band = $sheet.Shapes.AddShape(1, 200.0, $plotBottom - 150, 300.0, $bandBottom - ($plotBottom - 150))

    # The closing line, as the renderer would draw it at the anchor/pad boundary.
    $line = $sheet.Shapes.AddLine(200.0, $plotBottom + $anchorActual, 500.0, $plotBottom + $anchorActual)
    $line.Line.Weight = $frameWidthPt

    $beforeHeight = $band.Height
    $beforeAnchor = "$($band.TopLeftCell.Address($false, $false))/$($band.BottomRightCell.Address($false, $false))"

    $sheet.Rows.Item($InsertRow).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)

    $afterAnchor = "$($band.TopLeftCell.Address($false, $false))/$($band.BottomRightCell.Address($false, $false))"
    $heightDelta = [Math]::Round($band.Height - $beforeHeight, 3)
    $bandBottomAfter = $band.Top + $band.Height
    $lineBottom = $line.Top + $line.Height

    $outcome = if ([Math]::Abs($heightDelta - $bodyHeightPt) -lt 0.01) { 'STRETCHES' }
               elseif ([Math]::Abs($heightDelta) -lt 0.01) { 'SLIDES (bug)' }
               else { "delta=$heightDelta" }

    # Does the band still stop inside the line's coverage?
    $covered = ($bandBottomAfter -le ($lineBottom + 0.01)) -and ($bandBottomAfter -ge ($line.Top - 0.01))

    Write-Host ($headerFormat -f $InsertLabel, $anchorActual, $heightDelta, $outcome, "$beforeAnchor->$afterAnchor")
    Write-Host ("   band bottom {0}  line [{1}..{2}]  covered={3}" -f `
        [Math]::Round($bandBottomAfter, 3), [Math]::Round($line.Top, 3), [Math]::Round($lineBottom, 3), $covered)

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

    Write-Host '=== Q1: the row-height clamp (write then read back) ==='
    $book = $excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    Write-Host ($headerFormat -f 'Requested', 'ReadBack', 'Honoured', '', '')
    foreach ($candidate in @(0.25, 0.5, 0.6, 0.75, 1.0, 1.5, 2.0)) {
        $sheet.Rows.Item(1).RowHeight = $candidate
        $actual = $sheet.Rows.Item(1).RowHeight
        $honoured = if ([Math]::Abs($actual - $candidate) -lt 0.001) { 'yes' } else { 'NO - clamped' }
        Write-Host ($headerFormat -f $candidate, $actual, $honoured, '', '')
    }
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)

    Write-Host ''
    Write-Host '=== Q2: does the anchor-row layout stretch? ==='
    Write-Host 'Insert at row 9 is the ANCHOR row (one past the body) = the append target.'
    Write-Host ''

    foreach ($anchor in @(0.5, 1.0)) {
        Measure-Case -Excel $excel -AnchorHeightPt $anchor -InsertRow 9 -InsertLabel "anchor=$anchor row9"
    }
    # Control: an insert inside the body must also stretch.
    Measure-Case -Excel $excel -AnchorHeightPt 0.5 -InsertRow 5 -InsertLabel 'anchor=0.5 row5'

    Write-Host ''
    Write-Host 'PROBE COMPLETE'
}
finally {
    if ($excel) { $excel.Quit() }
    if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}