#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable spike: does an appended Gantt body row keep the Start|Finish rule, and in
    which form (direct formatting, table style, or neither)?

.DESCRIPTION
    Owner report (2026-11-15/16): "the right border of the start column and the left border
    of the finish column are not plotting when new entities are added; the primary
    initialisation row is correct, but every row after that is formatted incorrectly".

    The add-in writes cell borders in exactly one place: ExcelWorkbookInitialiser.
    ApplyBorderColour (mirrored by ExcelWorkbookRecovery), which runs during Initialise and
    Repair over the table's Range AS IT IS THEN. ListObjects.Add over a header-only range
    leaves the table with one blank body row, so that one row is the only body row that ever
    receives the theme's border write. An Add command then inserts a worksheet row with
    xlFormatFromRightOrBelow, so the new body row copies its cell format from the row BELOW
    it - which on the append branch is the reserved padding row, not a body row - and no
    later pass rewrites cell formatting (a Refresh deliberately writes none).

    This probe builds that exact sequence on a private, invisible Excel instance whose
    workbook is never saved, and measures the C|D and D|E boundaries on the header row, the
    initial body row and the appended body row:

      direct   Borders(xlEdgeRight)/(xlEdgeLeft) on each cell, and the pair's
               xlInsideVertical - what the cell itself carries.
      display  the same edges through each cell's own DisplayFormat, which is what the host
               would paint, with the colour it reports.
      fill     DisplayFormat.Interior.Color of both cells, so a rule painted in (or near)
               the fill's colour is visible as such rather than as "present".

    Three phases, because the reported correlation is provenance, and the two candidate
    causes differ in which phase moves the reading:

      A  production behaviour: initialise-shaped table, then append with
         xlFormatFromRightOrBelow (the value ExcelGanttRowInserter uses).
      B  the same, with the warning fill applied to the initial row's Start and Finish cells
         first - the state the owner's sheet is in when the rows are invalid.
      C  a fresh table appended with xlFormatFromLeftOrAbove, the default the add-in stopped
         using on 2026-11-09 to stop the pastel fill being cloned into the new row.

    It writes only to its own scratch workbook and never saves it. It never attaches to a
    running Excel, so it cannot touch the owner's workbook.

.NOTES
    Disposable spike in the style of probe-rowinsert-anchoring.ps1: results are printed as
    they are measured. It is not part of any gate. The measurement it produces belongs in
    docs/STATUS.md when a fix is chosen.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Excel's own values, as plain integers so the probe needs no interop assembly loaded.
# XlListObjectSourceType / XlYesNoGuess.
$xlSrcRange = 1
$xlYes = 1

# XlInsertShiftDirection / XlInsertFormatOrigin.
$xlShiftDown = -4121
$xlFormatFromLeftOrAbove = 0
$xlFormatFromRightOrBelow = 1

# XlBordersIndex.
$xlEdgeLeft = 7
$xlEdgeTop = 8
$xlEdgeBottom = 9
$xlEdgeRight = 10
$xlInsideVertical = 11

# XlBorderWeight.
$xlThin = 2

# XlLineStyle.
$xlLineStyleContinuous = 1

# The catalogue values the add-in writes: GanttTableAppearance.BorderColour reads
# MinorGridStroke (#D9D9D9), and the row warning highlighter fills with RGB(255,200,200)
# (OLE 0x00C8C8FF).
$themeBorderColour = 0xD9D9D9
$pastelRed = 0xC8C8FF

$headers = @('Type', 'Description', 'Start', 'Finish', 'Duration')

function Format-OleColour {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    $number = [int]$Value
    if ($number -lt 0) {
        $number += 0x1000000
    }

    $red = $number -band 0xFF
    $green = ($number -shr 8) -band 0xFF
    $blue = ($number -shr 16) -band 0xFF
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
}

# Builds the sheet the way Initialise does, as far as cell presentation is concerned: a
# header-only table (Excel adds the one blank body row), the neutral built-in style, no
# stripes or filters, then ApplyBorderColour's range-wide colour plus the four outside
# weights.
function Build-Table {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] [string] $OriginAddress,
        # Each phase needs its own name: two tables cannot share one, and the host refuses
        # the second Add with 0x800A03EC.
        [Parameter(Mandatory)] [string] $TableName
    )

    for ($index = 0; $index -lt $headers.Count; $index++) {
        $Sheet.Range($OriginAddress).Offset(0, $index).Value2 = $headers[$index]
    }

    $range = $Sheet.Range($OriginAddress).Resize(1, $headers.Count)
    $table = $Sheet.ListObjects.Add($xlSrcRange, $range, $null, $xlYes, $null)
    $table.Name = $TableName
    $table.ShowAutoFilter = $false
    $table.TableStyle = 'TableStyleLight1'
    $table.ShowTableStyleRowStripes = $false
    $table.ShowTableStyleColumnStripes = $false

    $borders = $table.Range.Borders
    $borders.Color = $themeBorderColour
    foreach ($edge in @($xlEdgeTop, $xlEdgeBottom, $xlEdgeLeft, $xlEdgeRight)) {
        $borders.Item($edge).Weight = $xlThin
    }

    return $table
}

# The append branch of ExcelGanttRowInserter: a real worksheet row inserted one past the
# table's last row (so the reserved row is displaced, not consumed), then ListRows.Add.
function Add-BodyRow {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] $Table,
        [Parameter(Mandatory)] [int] $CopyOrigin,
        [Parameter(Mandatory)] [double] $RowHeight
    )

    $tableRange = $Table.Range
    $lastRow = $tableRange.Row + $tableRange.Rows.Count - 1
    # Captured: Range.Insert returns a Variant, which would otherwise join this function's
    # return value and turn the row number into an array.
    $null = $Sheet.Rows.Item($lastRow + 1).Insert($xlShiftDown, $CopyOrigin)
    $null = $Table.ListRows.Add()
    $newRow = $lastRow + 1
    $Sheet.Rows.Item($newRow).RowHeight = $RowHeight

    # WriteRow: values only. No formatting is written, which is the point of the phase.
    $values = @('As-Planned Activity', 'probe row', $null, $null, $null)
    for ($index = 0; $index -lt $headers.Count; $index++) {
        if ($null -ne $values[$index]) {
            $Sheet.Cells.Item($newRow, $tableRange.Column + $index).Value2 = $values[$index]
        }
    }

    return $newRow
}

function Write-Readings {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] [string] $Phase,
        # object[], not int[]: a value that came out of a COM proxy can be an object
        # wrapper, and the cast to the row index belongs at the point of use.
        [Parameter(Mandatory)] [object[]] $Rows
    )

    # C|D and D|E: Description|Start is the control boundary, Start|Finish is the report.
    $boundaries = @(
        [pscustomobject]@{ Name = 'Description|Start'; Left = 3; Right = 4 },
        [pscustomobject]@{ Name = 'Start|Finish'; Left = 4; Right = 5 }
    )

    Write-Host ("=== {0} ===" -f $Phase)
    foreach ($boundary in $boundaries) {
        Write-Host ("--- {0} ---" -f $boundary.Name)
        foreach ($rowValue in $Rows) {
            $row = [int]$rowValue
            $leftCell = $Sheet.Cells.Item($row, $boundary.Left)
            $rightCell = $Sheet.Cells.Item($row, $boundary.Right)

            $leftDirect = [int]$leftCell.Borders.Item($xlEdgeRight).LineStyle
            $rightDirect = [int]$rightCell.Borders.Item($xlEdgeLeft).LineStyle
            $pairDirect = [int]$Sheet.Range($leftCell, $rightCell).Borders.Item($xlInsideVertical).LineStyle

            $leftDisplay = $leftCell.DisplayFormat.Borders.Item($xlEdgeRight)
            $rightDisplay = $rightCell.DisplayFormat.Borders.Item($xlEdgeLeft)

            $verdict = if ([int]$leftDisplay.LineStyle -eq $xlLineStyleContinuous -and [int]$rightDisplay.LineStyle -eq $xlLineStyleContinuous) {
                'PAINTED'
            }
            else {
                'NOT-PAINTED'
            }

            Write-Host ('  row {0,2} h={1,5} direct={2}/{3}/{4} display={5}({6})/{7}({8}) fill={9}/{10} => {11}' -f `
                    $row,
                    [Math]::Round([double]$Sheet.Rows.Item($row).RowHeight, 3),
                    $leftDirect, $rightDirect, $pairDirect,
                    [int]$leftDisplay.LineStyle, (Format-OleColour $leftDisplay.Color),
                    [int]$rightDisplay.LineStyle, (Format-OleColour $rightDisplay.Color),
                    (Format-OleColour $leftCell.DisplayFormat.Interior.Color),
                    (Format-OleColour $rightCell.DisplayFormat.Interior.Color),
                    $verdict)
        }

        Write-Host ''
    }
}

$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'Gantt'
    $sheet.Rows.Item(2).RowHeight = 18

    Write-Host ("Excel version: {0}" -f $excel.Version)
    Write-Host ''

    # --- A: append with the production CopyOrigin --------------------------------
    $table = Build-Table -Sheet $sheet -OriginAddress 'B3' -TableName 'tblGanttData'
    $initialRow = $table.Range.Row + 1
    $appendedRow = Add-BodyRow -Sheet $sheet -Table $table -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'A  append with xlFormatFromRightOrBelow (production)' -Rows @(3, $initialRow, $appendedRow)

    # --- B: the same, with the owner's pink warning fill on the initial row ------
    $table2 = Build-Table -Sheet $sheet -OriginAddress 'B12' -TableName 'tblProbeB'
    $initialRow2 = $table2.Range.Row + 1
    # Start is worksheet column D and Finish is E in this five-column layout.
    $Sheet.Cells.Item($initialRow2, 4).Interior.Color = $pastelRed
    $Sheet.Cells.Item($initialRow2, 5).Interior.Color = $pastelRed
    $appendedRow2 = Add-BodyRow -Sheet $sheet -Table $table2 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Host '  (phase B: the initial row''s Start/Finish cells carry the pastel warning fill)'
    Write-Host ('  appended row fill: Start={0} Finish={1}' -f `
            (Format-OleColour $sheet.Cells.Item($appendedRow2, 4).DisplayFormat.Interior.Color),
            (Format-OleColour $sheet.Cells.Item($appendedRow2, 5).DisplayFormat.Interior.Color))
    Write-Readings -Sheet $sheet -Phase 'B  append with xlFormatFromRightOrBelow, initial row pink' -Rows @(12, $initialRow2, $appendedRow2)

    # --- C: append with the default CopyOrigin the add-in stopped using ---------
    $table3 = Build-Table -Sheet $sheet -OriginAddress 'B21' -TableName 'tblProbeC'
    $initialRow3 = $table3.Range.Row + 1
    $appendedRow3 = Add-BodyRow -Sheet $sheet -Table $table3 -CopyOrigin $xlFormatFromLeftOrAbove -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'C  append with xlFormatFromLeftOrAbove (default)' -Rows @(21, $initialRow3, $appendedRow3)

    # --- D: which write removes the shared rule, and does re-applying the border restore it?
    # This isolates the only write in the pipeline that touches the two reported cells after
    # the table is presented: the validation report's pastel fill.
    $table4 = Build-Table -Sheet $sheet -OriginAddress 'B30' -TableName 'tblProbeD'
    $fillRow = [int]$table4.Range.Row + 1
    Write-Readings -Sheet $sheet -Phase 'D1 the table exactly as Initialise leaves it' -Rows @($fillRow)

    $Sheet.Cells.Item($fillRow, 4).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'D2 Start cell filled (the report fills one cell at a time)' -Rows @($fillRow)

    $Sheet.Cells.Item($fillRow, 5).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'D3 Start and Finish both filled' -Rows @($fillRow)

    $pair = $Sheet.Range($Sheet.Cells.Item($fillRow, 4), $Sheet.Cells.Item($fillRow, 5))
    $pairBorders = $pair.Borders
    $pairBorders.Color = $themeBorderColour
    Write-Readings -Sheet $sheet -Phase 'D4 the pair''s border colour re-applied after the fills' -Rows @($fillRow)

    $tableBorders = $table4.Range.Borders
    $tableBorders.Color = $themeBorderColour
    Write-Readings -Sheet $sheet -Phase 'D5 the table range''s border colour re-applied (ApplyBorderColour)' -Rows @($fillRow)

    $appended4 = Add-BodyRow -Sheet $sheet -Table $table4 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'E1 a body row appended after the repair' -Rows @($appended4)

    $Sheet.Rows.Item($appended4).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'E2 the whole-row fill the report also uses' -Rows @($appended4)

    # --- F: the owner's exact sequence. Two body rows, the LAST one filled by the report,
    # then one more Add. Phase B lost the shared rule somewhere in this sequence; this
    # isolates whether the fill, the append, or the pair of them is the trigger, with the
    # unfilled row above it as the control.
    $table5 = Build-Table -Sheet $sheet -OriginAddress 'B40' -TableName 'tblProbeF'
    $firstRow5 = Add-BodyRow -Sheet $sheet -Table $table5 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    $Sheet.Cells.Item($firstRow5, 4).Interior.Color = $pastelRed
    $Sheet.Cells.Item($firstRow5, 5).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'F1 one appended row, filled, before any further Add' -Rows @($firstRow5)

    $secondRow5 = Add-BodyRow -Sheet $sheet -Table $table5 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'F2 after ONE more Add (the control row above is unfilled)' -Rows @($firstRow5, $secondRow5)

    $thirdRow5 = Add-BodyRow -Sheet $sheet -Table $table5 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'F3 after a second further Add' -Rows @($firstRow5, $secondRow5, $thirdRow5)

    # --- G: trigger, repair, durability. Phase B lost the rule on the row that
    # ListObjects.Add created once a row was appended below it; phase F (fill on an
    # APPENDED row) did not. This fills the initial row, measures, appends, measures the
    # same row again, then re-applies the border colour and appends once more.
    $table6 = Build-Table -Sheet $sheet -OriginAddress 'B50' -TableName 'tblProbeG'
    $initialRow6 = [int]$table6.Range.Row + 1
    $Sheet.Cells.Item($initialRow6, 4).Interior.Color = $pastelRed
    $Sheet.Cells.Item($initialRow6, 5).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'G1 the initial row filled, before any Add' -Rows @($initialRow6)

    $afterG = Add-BodyRow -Sheet $sheet -Table $table6 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'G2 after ONE Add below the filled initial row' -Rows @($initialRow6, $afterG)

    $repairBorders = $table6.Range.Borders
    $repairBorders.Color = $themeBorderColour
    Write-Readings -Sheet $sheet -Phase 'G3 border colour re-applied over the table range' -Rows @($initialRow6, $afterG)

    $afterG2 = Add-BodyRow -Sheet $sheet -Table $table6 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'G4 after a further Add (does the repair survive?)' -Rows @($initialRow6, $afterG, $afterG2)

    # --- H: is the trigger BOTH cells of the pair being filled, or either one? The owner's
    # sheet shows row 4 with its boundary and the rows under it without, and row 4 is the row
    # whose Start cell alone is filled in that screenshot. If one filled cell is harmless and
    # two filled cells are not, the two observations are the same defect.
    $table7 = Build-Table -Sheet $sheet -OriginAddress 'B60' -TableName 'tblProbeH'
    $initialRow7 = [int]$table7.Range.Row + 1
    $Sheet.Cells.Item($initialRow7, 4).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'H1 only the Start cell filled' -Rows @($initialRow7)

    $afterH = Add-BodyRow -Sheet $sheet -Table $table7 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'H2 after an Add (only Start was filled)' -Rows @($initialRow7, $afterH)

    $Sheet.Cells.Item($initialRow7, 5).Interior.Color = $pastelRed
    Write-Readings -Sheet $sheet -Phase 'H3 the Finish cell filled as well' -Rows @($initialRow7, $afterH)

    $afterH2 = Add-BodyRow -Sheet $sheet -Table $table7 -CopyOrigin $xlFormatFromRightOrBelow -RowHeight 18
    Write-Readings -Sheet $sheet -Phase 'H4 after another Add (both cells were filled)' -Rows @($initialRow7, $afterH, $afterH2)

    # --- I: which write actually REMOVES a fill? ---------------------------------
    # The product clears with Interior.Color = xlColorIndexNone (-4142). Measure whether the
    # host honours it, and what the alternatives read back as.
    $table8 = Build-Table -Sheet $sheet -OriginAddress 'B70' -TableName 'tblProbeI'
    $fillRow8 = [int]$table8.Range.Row + 1
    $target8 = $Sheet.Rows.Item($fillRow8)
    $target8.Interior.Color = $pastelRed
    Write-Host ("=== I which write removes a fill ===")
    Write-Host ("  I1 filled with the report's red      : raw={0} colour={1}" -f ([int]$target8.Interior.Color), (Format-OleColour $target8.Interior.Color))
    $target8.Interior.Color = -4142
    Write-Host ("  I2 after Interior.Color = -4142      : raw={0} colour={1}" -f ([int]$target8.Interior.Color), (Format-OleColour $target8.Interior.Color))
    $target8.Interior.Color = $pastelRed
    $target8.Interior.ColorIndex = -4142
    Write-Host ("  I3 after Interior.ColorIndex = -4142 : raw={0} colour={1}" -f ([int]$target8.Interior.Color), (Format-OleColour $target8.Interior.Color))
    $target8.Interior.Color = $pastelRed
    $target8.Interior.Pattern = -4142
    Write-Host ("  I4 after Interior.Pattern = -4142    : raw={0} colour={1}" -f ([int]$target8.Interior.Color), (Format-OleColour $target8.Interior.Color))

    Write-Host 'PROBE COMPLETE (the scratch workbook is discarded, never saved)'
}
finally {
    $excel.ScreenUpdating = $true

    if ($null -ne $workbook) {
        $workbook.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
