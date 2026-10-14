#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: does a worksheet-row MOVE (an
    entire-row Cut to a destination) shift the moved
    rows' Gantt shapes, and what happens to the table,
    the intervening rows and the reserved padding row?

.DESCRIPTION
    R5.4 D1-amended implemented Move Up/Down as content
    rotation: two windows of cell content exchange via
    bulk writes, so the table entries reorder but the
    worksheet rows -- and the shapes anchored to them --
    stay put until the next Refresh. The owner's 2026-10-12
    direction: the move must move the ENTIRE row, not just
    the table entry. ADR-0036 Q2 measured a worksheet-row
    INSERT inside a ListObject range as accepted with every
    shape following, and the D10 probe measured the DELETE
    the same way; this probe measures the remaining single
    host operation a move needs, an entire-row Cut:

      Q1. Is Rows.Cut(destination) accepted inside a
          ListObject range, and does the content land at
          the destination (the swap/rotation semantics)?
      Q2. Do the shapes anchored to the moved rows and the
          intervening rows follow -- the shift an insert or
          delete produces -- or do they stay pinned to their
          anchor cells (a content write's behaviour)?
      Q3. Does the multi-row window case (a top-level
          group's window moving past its neighbour, applied
          as the inverse cut of the neighbour up) rotate the
          union span correctly?
      Q4. Does the ListObject stay valid (range and row
          count), and does the reserved 6pt padding row stay
          directly below the table?

    Both sections cut the LOWER window to the UPPER
    window's top row -- the one Cut shape that implements
    both move directions (a move-down of window W past
    neighbour N is the same host operation as a move-up of
    N past W).

    Disposable by design: it drives a throwaway workbook
    through COM and prints measurements. It is not a test,
    is not run by verify.ps1, and its findings are recorded
    as evidence in the R5.4 work item.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0

# XlPlacement values (PIA enum, printed raw because pure COM
# has no assembly to load the enum from): 1 = xlMove,
# 2 = xlMoveAndSize, 3 = xlFreeFloating.
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

function Show-Table {
    param($Table, $Sheet, [string] $Label, [double] $PaddingHeightPt)

    $last = $Table.Range.Row + $Table.Range.Rows.Count - 1
    Write-Host ("  {0,-7} table={1}..{2} ListRows={3} range={4} rowBelowHeight={5}pt isPadding={6}" -f `
        $Label, $Table.Range.Row, $last, $Table.ListRows.Count, `
        $Table.Range.Address($false, $false), $Sheet.Rows.Item($last + 1).RowHeight, `
        ($Sheet.Rows.Item($last + 1).RowHeight -eq $PaddingHeightPt))
}

function New-MoveSheet {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($Workbook, [string] $Name, [int] $BodyRowCount)

    if (-not $PSCmdlet.ShouldProcess($Name, 'Create worksheet')) {
        return
    }

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $Name
    $lastBodyRow = 1 + $BodyRowCount
    $table = $sheet.ListObjects.Add(1, $sheet.Range("A1:C{0}" -f $lastBodyRow), $null, 1)
    $table.Name = 'tbl' + $Name
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..$lastBodyRow) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item($lastBodyRow + 1).RowHeight = $paddingHeightPt
    return $sheet
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host

    $workbook = $excel.Workbooks.Add()

    # ---- Section 1 (Q1, Q2, Q4): a single-row move-up ----
    # Five body rows A..E (rows 2..6). Moving C (row 4) up
    # past B (row 3) is Rows.Item(4).Cut(Rows.Item(3)); the
    # expected content is A, C, B, D, E.
    $single = New-MoveSheet $workbook 'Single' 5
    $shapes = $single.Shapes
    $tops = @(20.0, 38.0, 56.0, 74.0, 92.0)
    $names = @('barRow2', 'barRow3', 'barRow4', 'barRow5', 'barRow6')
    for ($i = 0; $i -lt 5; $i++) {
        $null = $shapes.AddShape(1, 200.0, $tops[$i], 80.0, 8.0).Name = $names[$i]
        $single.Range('A{0}' -f ($i + 2)).Value2 = [string]([char](65 + $i))
    }

    $before = @{}
    Write-Host '=== Section 1: single-row move-up, C (row 4) past B (row 3) ==='
    foreach ($name in $names) {
        $shape = $shapes.Item($name)
        $before[$name] = $shape.Top
        Show-Placement $shape $name
        Show-Shape $shape $name $before 'before'
    }
    Show-Table $single.ListObjects.Item('tblSingle') $single 'before' $paddingHeightPt

    try {
        $single.Rows.Item(4).Cut($single.Rows.Item(3))
        Write-Host '  Rows.Item(4).Cut(Rows.Item(3)): accepted'
    }
    catch {
        Write-Host ("  Rows.Item(4).Cut(Rows.Item(3)): REFUSED -- {0}" -f $_.Exception.Message)
        Write-Host '  PROBE COMPLETE (cut refused; no further measurement possible)'
        return
    }

    Write-Host ("  content: A2={0} A3={1} A4={2} A5={3} A6={4} (expected A C B D E)" -f `
        $single.Range('A2').Value2, $single.Range('A3').Value2, $single.Range('A4').Value2, `
        $single.Range('A5').Value2, $single.Range('A6').Value2)
    Show-Table $single.ListObjects.Item('tblSingle') $single 'after ' $paddingHeightPt
    foreach ($name in $names) {
        Show-Shape $shapes.Item($name) $name $before 'after '
    }
    Write-Host ("  shapes remaining: Count={0}" -f $shapes.Count)

    # ---- Section 2 (Q3, Q4): the group-window move-down ----
    # Five body rows: A with children A1, A2 (rows 2..4),
    # top-level B (row 5), top-level E (row 6). Moving A's
    # window down past B is applied as the inverse cut: the
    # lower window (B, row 5) up to the upper window's top
    # (row 2), Rows.Item(5).Cut(Rows.Item(2)); the expected
    # content is B, A, A1, A2, E.
    $window = New-MoveSheet $workbook 'Window' 5
    $wShapes = $window.Shapes
    for ($i = 0; $i -lt 5; $i++) {
        $null = $wShapes.AddShape(1, 200.0, $tops[$i], 80.0, 8.0).Name = $names[$i]
    }
    $window.Range('A2').Value2 = 'A'
    $window.Range('A3').Value2 = 'A1'
    $window.Range('A4').Value2 = 'A2'
    $window.Range('A5').Value2 = 'B'
    $window.Range('A6').Value2 = 'E'

    $wBefore = @{}
    Write-Host ''
    Write-Host '=== Section 2: group-window move-down, A+A1+A2 (rows 2-4) past B (row 5) ==='
    Write-Host '  (applied as the inverse cut: Rows.Item(5).Cut(Rows.Item(2)))'
    foreach ($name in $names) {
        $shape = $wShapes.Item($name)
        $wBefore[$name] = $shape.Top
        Show-Shape $shape $name $wBefore 'before'
    }
    Show-Table $window.ListObjects.Item('tblWindow') $window 'before' $paddingHeightPt

    try {
        $window.Rows.Item(5).Cut($window.Rows.Item(2))
        Write-Host '  Rows.Item(5).Cut(Rows.Item(2)): accepted'
    }
    catch {
        Write-Host ("  Rows.Item(5).Cut(Rows.Item(2)): REFUSED -- {0}" -f $_.Exception.Message)
        Write-Host '  PROBE COMPLETE (cut refused; no further measurement possible)'
        return
    }

    Write-Host ("  content: A2={0} A3={1} A4={2} A5={3} A6={4} (expected B A A1 A2 E)" -f `
        $window.Range('A2').Value2, $window.Range('A3').Value2, $window.Range('A4').Value2, `
        $window.Range('A5').Value2, $window.Range('A6').Value2)
    Show-Table $window.ListObjects.Item('tblWindow') $window 'after ' $paddingHeightPt
    foreach ($name in $names) {
        Show-Shape $wShapes.Item($name) $name $wBefore 'after '
    }
    Write-Host ("  shapes remaining: Count={0}" -f $wShapes.Count)

    # ---- Section 3 (Q5): the insert+delete rotation ----
    # The only mechanism that can rotate rows inside a
    # ListObject (a Cut is a destructive overwrite, Q1).
    # Five body rows A..E (rows 2..6). Moving C (row 4)
    # up past B (row 3): insert a row at row 3 (the host
    # shifts B..E down and the ListObject absorbs the row),
    # write C's saved content into the inserted row, then
    # delete the old C row (now row 5, shifted down by the
    # insert). Expected content: A, C, B, D, E. The open
    # question is the moved row's own shape: the delete
    # host-deletes the shape anchored in the deleted row
    # (the D10 measurement), and the insert only SHIFTS
    # shapes -- so C's bar is predicted to vanish with the
    # old row while every other shape lands correctly.
    # XlInsertShiftDirection.xlShiftDown and
    # XlInsertFormatOrigin.xlFormatFromLeftOrAbove as raw
    # values (pure COM has no assembly to load the enum
    # from) -- the same pair the production inserter
    # passes (ExcelGanttRowInserter.InsertWorksheetRow).
    $xlShiftDown = -4121
    $xlFormatFromLeftOrAbove = 0

    $rotate = New-MoveSheet $workbook 'Rotate' 5
    $rShapes = $rotate.Shapes
    for ($i = 0; $i -lt 5; $i++) {
        $null = $rShapes.AddShape(1, 200.0, $tops[$i], 80.0, 8.0).Name = $names[$i]
        $rotate.Range('A{0}' -f ($i + 2)).Value2 = [string]([char](65 + $i))
    }

    $rBefore = @{}
    Write-Host ''
    Write-Host '=== Section 3: rotation by insert+delete, C (row 4) up past B (row 3) ==='
    foreach ($name in $names) {
        $shape = $rShapes.Item($name)
        $rBefore[$name] = $shape.Top
        Show-Shape $shape $name $rBefore 'before'
    }
    Show-Table $rotate.ListObjects.Item('tblRotate') $rotate 'before' $paddingHeightPt

    try {
        $rotate.Rows.Item(3).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
        Write-Host '  Rows.Item(3).Insert(xlShiftDown, xlFormatFromLeftOrAbove): accepted'
    }
    catch {
        Write-Host ("  Rows.Item(3).Insert(...): REFUSED -- {0}" -f $_.Exception.Message)
        Write-Host '  PROBE COMPLETE (insert refused; no further measurement possible)'
        return
    }

    # The moved row's saved content, written into the
    # inserted row -- what the mover adapter would do with
    # the body matrix it already holds.
    $rotate.Range('A3').Value2 = 'C'

    # The old C row: row 4, displaced to row 5 by the insert.
    try {
        $rotate.Rows.Item(5).Delete()
        Write-Host '  Rows.Item(5).Delete() (the displaced old C row): accepted'
    }
    catch {
        Write-Host ("  Rows.Item(5).Delete(): REFUSED -- {0}" -f $_.Exception.Message)
        Write-Host '  PROBE COMPLETE (delete refused; the table holds a duplicated row)'
        return
    }

    Write-Host ("  content: A2={0} A3={1} A4={2} A5={3} A6={4} (expected A C B D E)" -f `
        $rotate.Range('A2').Value2, $rotate.Range('A3').Value2, $rotate.Range('A4').Value2, `
        $rotate.Range('A5').Value2, $rotate.Range('A6').Value2)
    Show-Table $rotate.ListObjects.Item('tblRotate') $rotate 'after ' $paddingHeightPt
    Write-Host ("  heights: r2={0} r3={1} r4={2} r5={3} r6={4} r7={5} r8={6}" -f `
        $rotate.Rows.Item(2).RowHeight, $rotate.Rows.Item(3).RowHeight, `
        $rotate.Rows.Item(4).RowHeight, $rotate.Rows.Item(5).RowHeight, `
        $rotate.Rows.Item(6).RowHeight, $rotate.Rows.Item(7).RowHeight, `
        $rotate.Rows.Item(8).RowHeight)
    foreach ($name in $names) {
        try {
            Show-Shape $rShapes.Item($name) $name $rBefore 'after '
        }
        catch {
            Write-Host ("  {0,-8} after  GONE -- the host removed it with the deleted source row" -f $name)
        }
    }
    Write-Host ("  shapes remaining: Count={0} (expected 4 -- the moved row's bar is host-deleted)" -f $rShapes.Count)

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