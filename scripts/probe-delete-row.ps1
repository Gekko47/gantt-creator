#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe for the R5.4 Remove command: what does a live row DELETE do?

.DESCRIPTION
    Move up/down is proven (ExcelParentChildMutatorTests); Remove is not.
    "Delete the row" is the one content-mutating operation whose COM
    behaviour this repository has never measured, and it differs from an
    insert in the ways that matter most: a ListRows.Delete removes the row
    from the ListObject, the shapes above the deleted row keep their
    anchors, the shapes below move, and the reserved padding row below the
    table (ADR-0035/0036) has to survive or the next Refresh breaks.

    Five questions:
      Q1. Does ListRows.Delete(index) shrink the ListObject by exactly one
          row, and does the row below the table (the padding row) survive?
      Q2. What happens to a shape anchored to a row BELOW the deleted row,
          and to a shape anchored ABOVE it?
      Q3. Does deleting the LAST body row leave an empty but valid
          ListObject (ListRows.Count == 0) rather than deleting the table?
      Q4. Does the reserved padding row keep its height (6pt) and stay
          directly below the table after the delete?
      Q5. Does the body-row height (18pt) of the surviving rows change?

    Disposable by design: it drives a throwaway workbook through COM and
    prints measurements. It is not a test, is not run by verify.ps1, and
    its findings are recorded as evidence in the Remove work item.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
    Write-Host ("  {0}: table={1}..{2} ListRows={3} rowBelow={4}pt rowBelowHeightPreserved={5}" -f `
        $Label, $Table.Range.Row, $last, $Table.ListRows.Count, $Sheet.Rows.Item($last + 1).RowHeight, ($Sheet.Rows.Item($last + 1).RowHeight -eq $paddingHeightPt))
}

function Show-Shapes {
    param($Shapes, [hashtable] $Before, [string] $Label)

    foreach ($name in @('barLow', 'barMid', 'barHigh')) {
        $shape = $Shapes.Item($name)
        Write-Host ("  {0,-8} {1,-6} Top {2,-7} -> {3,-7} delta={4,-7} TopLeftCell={5}" -f `
            $name, $Label, $Before[$name], $shape.Top, ($shape.Top - $Before[$name]), $shape.TopLeftCell.Address($false, $false))
    }
}

function Show-BodyHeights {
    param($Sheet, [int] $First, [int] $Count, [string] $Label)

    $hts = @()
    foreach ($r in ($First)..($First + $Count - 1)) { $hts += [string]([int]$Sheet.Rows.Item($r).RowHeight) }
    Write-Host ("  {0} body row heights (pt): {1}" -f $Label, ($hts -join ','))
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
    Write-Host '=== Q1/Q2/Q4/Q5: ListRows.Delete(index) -- table, shapes, padding row, heights ==='
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
    Show-BodyHeights $sheetA 2 3 'before'

    # Delete the MIDDLE body row (body index 2 of 3 = worksheet row 3).
    # barLow sits above it, barMid/barHigh below it.
    $tableA.ListRows.Item(2).Delete()
    Write-Host '  -- after ListRows.Item(2).Delete() --'
    Show-State $sheetA $tableA 'after '
    Show-Shapes $shapesA $beforeA 'after '
    Show-BodyHeights $sheetA 2 2 'after '

    # ---------------------------------------------------------------- Q3
    Write-Host ''
    Write-Host '=== Q3: does deleting the LAST body row leave an empty but valid table? ==='
    $d = Build-Sheet $workbook 'Q3'
    $sheetD = $d.Sheet; $tableD = $d.Table
    $shapesD = $sheetD.Shapes
    $null = $shapesD.AddShape(1, 200.0, 40.0, 80.0, 8.0).Name = 'barLow'
    $beforeD = @{}
    foreach ($n in @('barLow')) { $beforeD[$n] = $shapesD.Item($n).Top }
    Show-State $sheetD $tableD 'before'

    $tableD.ListRows.Item(3).Delete()
    Write-Host '  -- after ListRows.Item(3).Delete() (last body row) --'
    Show-State $sheetD $tableD 'after '
    $s = $shapesD.Item('barLow')
    Write-Host ("  barLow delta={0} TopLeftCell={1}" -f ($s.Top - $beforeD['barLow']), $s.TopLeftCell.Address($false, $false))

    # Can the empty table still take a row (the post-Remove re-add path)?
    try {
        $tableD.ListRows.Add(1)
        Write-Host "  ListRows.Add(1) on the emptied table: OK, ListRows.Count=$($tableD.ListRows.Count)"
    }
    catch {
        Write-Host "  ERR adding a row to the emptied table: $($_.Exception.Message)"
    }

    # ------------------------------------------------------------- Q3b
    # The empty-table edge Remove can actually reach: a one-row table whose
    # only body row is deleted. Q3 leaves two rows behind, so it never
    # measures the zero-data-row ListObject.
    Write-Host ''
    Write-Host '=== Q3b: deleting the ONLY body row -- does the ListObject survive at zero rows? ==='
    $e = $workbook.Worksheets.Add()
    $e.Name = 'Q3b'
    $tableE = $e.ListObjects.Add(1, $e.Range('A1:C2'), $null, 1)
    $tableE.Name = 'tblQ3b'
    $e.Rows.Item(1).RowHeight = $headerHeightPt
    $e.Rows.Item(2).RowHeight = $bodyHeightPt
    $e.Rows.Item(3).RowHeight = $paddingHeightPt
    Write-Host ("  before: ListRows={0} range={1}" -f $tableE.ListRows.Count, $tableE.Range.Address($false, $false))

    try {
        $tableE.ListRows.Item(1).Delete()
        $e2 = $tableE.Range.Address($false, $false)
        # A fresh ListObject auto-names its first header 'Column1'; the point of
        # the check is that the header row still holds a header after the body
        # row is gone, not which name Excel chose.
        $header = [string]$e.Range('A1').Value2
        Write-Host ("  after : ListRows={0} range={1} headerStillAtA1={2} (value='{3}')" -f `
            $tableE.ListRows.Count, $e2, (-not [string]::IsNullOrWhiteSpace($header)), $header)
    }
    catch {
        Write-Host "  ERR deleting the only body row: $($_.Exception.Message)"
    }

    try {
        $tableE.ListRows.Add(1) | Out-Null
        Write-Host ("  re-add after emptying: OK, ListRows={0} range={1}" -f $tableE.ListRows.Count, $tableE.Range.Address($false, $false))
    }
    catch {
        Write-Host "  ERR re-adding after emptying: $($_.Exception.Message)"
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