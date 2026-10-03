#Requires -Version 7.0
<#
.SYNOPSIS
    Second disposable probe: does ListRows.Add(POSITION) shift the worksheet at all?

.DESCRIPTION
    The first probe (probe-rowinsert-anchoring.ps1) proved a real worksheet row
    insert moves shapes, and fixed the APPEND branch. It never tested the
    POSITIONAL branch -- ListRows.Add(position) -- which is the branch taken when
    the active cell is inside the table. A live report after the fix shows that
    branch still does not move the Gantt shapes, while Excel's own
    "Insert Row" does.

    Three questions:
      Q1. Does ListRows.Add(position) move a shape anchored BELOW the insertion
          point, and does it move the row below the table (the padding row)?
      Q2. Does a plain worksheet Rows.Insert INSIDE a table's range auto-expand the
          ListObject into the new row -- so no ListRows.Add is needed at all?
      Q3. If so, is the inserted row a real data row that accepts a written value,
          and is it at the body height we want?
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = -4142

$bodyHeightPt = 18.0
$paddingHeightPt = 6.0
$headerHeightPt = 16.0

function Build-Sheet {
    param($Workbook, [string] $Name)

    $sheet = $Workbook.Worksheets.Add()
    $sheet.Name = $Name
    # header row 1 + 3 body rows (2,3,4); row 5 is the reserved padding row
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $table.Name = "tbl$Name"
    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..4) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }
    $sheet.Rows.Item(5).RowHeight = $paddingHeightPt
    return @{ Sheet = $sheet; Table = $table }
}

function Show-State {
    param($Sheet, $Table, [string] $Label)

    $last = $Table.Range.Row + $Table.Range.Rows.Count - 1
    Write-Host ("  {0}: table={1}..{2} ListRows={3} rowBelow={4}pt" -f `
        $Label, $Table.Range.Row, $last, $Table.ListRows.Count, $Sheet.Rows.Item($last + 1).RowHeight)
}

function Show-Shapes {
    param($Shapes, [hashtable] $Before, [string] $Label)

    foreach ($name in @('barLow', 'barMid', 'barHigh')) {
        $shape = $Shapes.Item($name)
        Write-Host ("  {0,-6} {1,-8} Top {2,-7} -> {3,-7} delta={4,-6} TopLeftCell={5}" -f `
            $name, $Label, $Before[$name], $shape.Top, ($shape.Top - $Before[$name]), $shape.TopLeftCell.Address($false, $false))
    }
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

    # ---------------------------------------------------------------- Q1
    Write-Host '=== Q1: does ListRows.Add(POSITION) move the sheet? ==='
    $a = Build-Sheet $workbook 'Q1'
    $sheetA = $a.Sheet; $tableA = $a.Table
    $shapesA = $sheetA.Shapes
    $null = $shapesA.AddShape(1, 200.0, 40.0, 80.0, 8.0).Name = 'barLow'
    $null = $shapesA.AddShape(1, 200.0, 70.0, 80.0, 8.0).Name = 'barMid'
    $null = $shapesA.AddShape(1, 200.0, 100.0, 80.0, 8.0).Name = 'barHigh'

    $beforeA = @{}
    foreach ($n in @('barLow', 'barMid', 'barHigh')) { $beforeA[$n] = $shapesA.Item($n).Top }
    Show-State $sheetA $tableA 'before'
    Show-Shapes $shapesA $beforeA 'before'

    # Active cell inside the table at body row 2 -> insert at body index 2,
    # which is what GetInsertionPosition computes.
    $null = $tableA.ListRows.Add(2)
    Write-Host '  -- after ListRows.Add(2) --'
    Show-State $sheetA $tableA 'after '
    Show-Shapes $shapesA $beforeA 'after '
    Write-Host ("  row 5 height now: {0}pt   row 6 height now: {1}pt" -f $sheetA.Rows.Item(5).RowHeight, $sheetA.Rows.Item(6).RowHeight)

    # ---------------------------------------------------------------- Q2
    Write-Host ''
    Write-Host '=== Q2: does a worksheet Rows.Insert INSIDE the table auto-expand it? ==='
    $b = Build-Sheet $workbook 'Q2'
    $sheetB = $b.Sheet; $tableB = $b.Table
    $shapesB = $sheetB.Shapes
    $null = $shapesB.AddShape(1, 200.0, 40.0, 80.0, 8.0).Name = 'barLow'
    $null = $shapesB.AddShape(1, 200.0, 70.0, 80.0, 8.0).Name = 'barMid'
    $null = $shapesB.AddShape(1, 200.0, 100.0, 80.0, 8.0).Name = 'barHigh'
    $beforeB = @{}
    foreach ($n in @('barLow', 'barMid', 'barHigh')) { $beforeB[$n] = $shapesB.Item($n).Top }

    Show-State $sheetB $tableB 'before'
    Show-Shapes $shapesB $beforeB 'before'

    # Insert a genuine worksheet row at row 3, INSIDE the table's range.
    $sheetB.Rows.Item(3).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    Write-Host '  -- after Rows(3).Insert(xlShiftDown) inside the table --'
    Show-State $sheetB $tableB 'after '
    Show-Shapes $shapesB $beforeB 'after '
    Write-Host ("  new row 3 height: {0}pt" -f $sheetB.Rows.Item(3).RowHeight)

    # ---------------------------------------------------------------- Q3
    Write-Host ''
    Write-Host '=== Q3: is the auto-expanded row a writable data row? ==='
    try {
        $listRow = $tableB.ListRows.Item(3)
        Write-Host "  ListRows.Item(3).Range.Address = $($listRow.Range.Address($false, $false))"
        $sheetB.Range('A3').Value2 = 'PROBE-ID'
        $sheetB.Range('B3').Value2 = 'PROBE-TYPE'
        Write-Host "  A3 now = '$($sheetB.Range('A3').Value2)'  B3 now = '$($sheetB.Range('B3').Value2)'"
        Write-Host "  A3 is inside the ListObject: $($null -ne $listRow)"
    }
    catch {
        Write-Host "  ERR: $($_.Exception.Message)"
    }

    Write-Host ''
    Write-Host '=== Q4: does a row insert directly BELOW the table auto-expand it? ==='
    $c = Build-Sheet $workbook 'Q4'
    $sheetC = $c.Sheet; $tableC = $c.Table
    $shapesC = $sheetC.Shapes
    $null = $shapesC.AddShape(1, 200.0, 40.0, 80.0, 8.0).Name = 'barLow'
    $null = $shapesC.AddShape(1, 200.0, 70.0, 80.0, 8.0).Name = 'barMid'
    $beforeC = @{}
    foreach ($n in @('barLow', 'barMid')) { $beforeC[$n] = $shapesC.Item($n).Top }

    Show-State $sheetC $tableC 'before'
    # Insert at one past the table's last row -- the APPEND position.
    $lastC = $tableC.Range.Row + $tableC.Range.Rows.Count - 1
    $sheetC.Rows.Item($lastC + 1).Insert($xlShiftDown, $xlFormatFromLeftOrAbove)
    Write-Host "  -- after Rows($($lastC + 1)).Insert, directly below the table --"
    Show-State $sheetC $tableC 'after '
    Write-Host ("  row {0} (new) height: {1}pt   row {2} (padding) height: {3}pt" -f `
        ($lastC + 1), $sheetC.Rows.Item($lastC + 1).RowHeight, ($lastC + 2), $sheetC.Rows.Item($lastC + 2).RowHeight)
    foreach ($n in @('barLow', 'barMid')) {
        $s = $shapesC.Item($n)
        Write-Host ("  {0,-7} delta={1}" -f $n, ($s.Top - $beforeC[$n]))
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