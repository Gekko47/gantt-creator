#Requires -Version 7.0
<#
.SYNOPSIS
    Read-only probe: does every visible table column boundary carry a PAINTED rule on
    every row, and is the Start|Finish boundary different from the rest?

.DESCRIPTION
    Owner report (2026-11-15/16): "the right border of the start column and the left border
    of the finish column are not plotting when new entities are added; the primary
    initialisation row is correct, but every row after that is formatted incorrectly".

    The end-to-end path that decides those cell borders is short, and only two places write
    to it:

      * Initialise and Repair call GanttTableAppearance through
        ExcelWorkbookInitialiser.ApplyBorderColour, which sets Borders.Color on
        tblGanttData's Range and gives the four outside edges xlThin. The table at that
        moment is the header row plus the single blank body row ListObjects.Add created,
        so that is the only body row that ever receives the theme's border colour directly.
      * Every later Add command inserts a worksheet row with
        xlFormatFromRightOrBelow (ExcelGanttRowInserter), so the new body row copies its
        cell format from the row BELOW it. For the append branch that source is the
        reserved padding row, not the last body row.
      * A Refresh writes no cell formatting at all (ADR-0060), so nothing ever revisits
        the appended rows.

    That leaves four possible states for one boundary on one row, and this probe reports
    which one the sheet is actually in:

      DIRECT      direct LineStyle on both edges => the cell carries the rule itself.
      STYLE-ONLY  no direct rule, but DisplayFormat reports a continuous rule => the table
                  style supplies it and the host should paint it.
      MISSING     neither reading reports a rule => nothing would draw that boundary.
      MERGED      one of the two cells is merged, so the boundary cannot exist.

    DisplayFormat is read PER CELL EDGE, not over the two-cell range: a multi-cell range's
    DisplayFormat is ambiguous about the interior rule (recorded in the reverted 38641fc
    notes), and ambiguity is what let an earlier green gate coexist with a visible defect.
    The colour is reported as well as the line style, because a rule painted in the fill's
    own colour is present in every reading and invisible on screen - the one class that a
    LineStyle-only assertion cannot see.

    It never writes to the workbook. With -Path it opens a private, invisible, read-only
    Excel instance and closes it again; without -Path it attaches to the Excel that is
    already running and leaves it running.

.PARAMETER Path
    Full path to the workbook to inspect. Opened read-only in a new, invisible Excel
    instance that is closed before the script returns.

.PARAMETER SheetName
    The Gantt worksheet to inspect. Defaults to the first worksheet that carries a table
    named tblGanttData, which is the contract's own resolution.

.EXAMPLE
    pwsh ./scripts/probe-table-grid-borders.ps1 -Path 'C:\Users\rossj\Documents\gantt.xlsx'

.EXAMPLE
    # The Gantt workbook is already open: attach to that Excel and read it.
    pwsh ./scripts/probe-table-grid-borders.ps1

.NOTES
    Disposable spike in the style of probe-plot-origin.ps1: results are printed as they are
    measured, because values built from COM proxies do not survive a PowerShell function
    boundary reliably. It is not part of any gate.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $Path,

    [Parameter(Mandatory = $false)]
    [string] $SheetName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dataTableName = 'tblGanttData'

# Excel's own values, as plain integers so the probe needs no interop assembly loaded.
# XlBordersIndex.
$xlEdgeLeft = 7
$xlEdgeRight = 10
$xlInsideVertical = 11

# XlLineStyle.
$xlLineStyleContinuous = 1

function Format-Number {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value) {
        return 'null'
    }

    return [Math]::Round([double]$Value, 3).ToString('0.###')
}

# Excel reports OLE_COLOR, which is BGR: 0x00BBGGRR. A rule whose colour equals the cell's
# fill is present in every reading and invisible to the eye, which is the whole reason the
# probe prints colour and not only LineStyle.
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

function Find-GanttSheet {
    param([Parameter(Mandatory)] $Workbook)

    $count = $Workbook.Worksheets.Count
    for ($index = 1; $index -le $count; $index++) {
        $sheet = $Workbook.Worksheets.Item($index)
        if ($null -ne (Find-GanttTable -Sheet $sheet)) {
            return $sheet
        }
    }

    return $null
}

function Find-GanttTable {
    param([Parameter(Mandatory)] $Sheet)

    $tables = $Sheet.ListObjects
    $tableCount = $tables.Count
    for ($tableIndex = 1; $tableIndex -le $tableCount; $tableIndex++) {
        $table = $tables.Item($tableIndex)
        if ($table.Name -eq $dataTableName) {
            return $table
        }
    }

    return $null
}

function Format-TableStyleName {
    param([Parameter(Mandatory)] $Table)

    # ListObject.TableStyle is object-typed: it reports either a TableStyle proxy or the
    # name it was assigned.
    $style = $Table.TableStyle
    if ($null -eq $style) {
        return 'null'
    }

    if ($style -is [string]) {
        return $style
    }

    return [string]$style.Name
}

function Write-TableFacts {
    param(
        [Parameter(Mandatory)] $Table,
        [Parameter(Mandatory)] $Sheet
    )

    Write-Host '=== 1. The table, its style, and the visible column map ==='
    Write-Host ("Table               : {0}" -f $Table.Name)
    Write-Host ("Table style         : {0}" -f (Format-TableStyleName -Table $Table))
    Write-Host ("Row stripes         : {0}" -f $Table.ShowTableStyleRowStripes)
    Write-Host ("Column stripes      : {0}" -f $Table.ShowTableStyleColumnStripes)
    Write-Host ("Range               : {0}" -f $Table.Range.Address())

    $body = $Table.DataBodyRange
    if ($null -eq $body) {
        Write-Host 'Data body           : (none)'
    }
    else {
        Write-Host ("Data body           : {0} ({1} row(s))" -f $body.Address(), $Table.ListRows.Count)
    }

    # The range-level inside rule is what ApplyBorderColour wrote at Initialise. It covers
    # the table as it was THEN, so its colour is the one the initialisation row carries and
    # the appended rows may not.
    $rangeBorders = $Table.Range.Borders
    $rangeInside = $rangeBorders.Item($xlInsideVertical)
    Write-Host ("Range inside-vert    : LineStyle={0} Weight={1} Colour={2}" -f `
            $rangeInside.LineStyle, $rangeInside.Weight, (Format-OleColour $rangeInside.Color))
    Write-Host ''

    Write-Host 'Visible columns (worksheet index, header, width, hidden):'
    $columnCount = $Table.ListColumns.Count
    $firstColumn = $Table.Range.Column
    for ($offset = 0; $offset -lt $columnCount; $offset++) {
        $column = $firstColumn + $offset
        $listColumn = $Table.ListColumns.Item($offset + 1)
        $worksheetColumn = $Sheet.Columns.Item($column)
        Write-Host ('  idx {0,2}  {1,-18} width={2,8} hidden={3}' -f `
                $column, $listColumn.Name, (Format-Number $worksheetColumn.ColumnWidth), $worksheetColumn.Hidden)
    }

    Write-Host ''
}

function Write-BoundaryMatrix {
    param(
        [Parameter(Mandatory)] $Table,
        [Parameter(Mandatory)] $Sheet
    )

    Write-Host '=== 2. Every visible boundary, on the header row and every body row ==='
    Write-Host 'direct  = Borders(xlEdgeRight) of the left cell / Borders(xlEdgeLeft) of the right cell / pair inside-vertical'
    Write-Host 'display = the same edges read through each cell''s own DisplayFormat, with the colour the host reports'
    Write-Host 'present = display says CONTINUOUS, which is what the host would paint'
    Write-Host ''

    $tableRange = $Table.Range
    $firstRow = $tableRange.Row
    $lastRow = $firstRow + $tableRange.Rows.Count - 1
    $firstColumn = $tableRange.Column
    $columnCount = $Table.ListColumns.Count

    $visible = @()
    for ($offset = 0; $offset -lt $columnCount; $offset++) {
        $column = $firstColumn + $offset
        $worksheetColumn = $Sheet.Columns.Item($column)
        if ([bool]$worksheetColumn.Hidden) {
            continue
        }

        $visible += [pscustomobject]@{
            Header  = $Table.ListColumns.Item($offset + 1).Name
            Column  = $column
            Width   = [double]$worksheetColumn.ColumnWidth
        }
    }

    Write-Host ("visible boundaries: {0}" -f (($visible | ForEach-Object { $_.Header }) -join ', '))
    Write-Host ''

    $findings = @()
    for ($boundary = 0; $boundary + 1 -lt $visible.Count; $boundary++) {
        $leftColumn = $visible[$boundary].Column
        $rightColumn = $visible[$boundary + 1].Column
        $name = '{0}|{1}' -f $visible[$boundary].Header, $visible[$boundary + 1].Header

        Write-Host ("--- {0} (worksheet columns {1} and {2}) ---" -f $name, $leftColumn, $rightColumn)
        for ($row = $firstRow; $row -le $lastRow; $row++) {
            $leftCell = $Sheet.Cells.Item($row, $leftColumn)
            $rightCell = $Sheet.Cells.Item($row, $rightColumn)

            $leftDirect = [int]$leftCell.Borders.Item($xlEdgeRight).LineStyle
            $rightDirect = [int]$rightCell.Borders.Item($xlEdgeLeft).LineStyle
            $pairDirect = [int]$Sheet.Range($leftCell, $rightCell).Borders.Item($xlInsideVertical).LineStyle

            $leftDisplay = $leftCell.DisplayFormat.Borders.Item($xlEdgeRight)
            $rightDisplay = $rightCell.DisplayFormat.Borders.Item($xlEdgeLeft)
            $leftDisplayStyle = [int]$leftDisplay.LineStyle
            $rightDisplayStyle = [int]$rightDisplay.LineStyle

            $merged = [bool]$leftCell.MergeCells -or [bool]$rightCell.MergeCells
            $leftFill = Format-OleColour $leftCell.DisplayFormat.Interior.Color
            $rightFill = Format-OleColour $rightCell.DisplayFormat.Interior.Color

            $verdict = if ($merged) {
                'MERGED'
            }
            elseif ($leftDisplayStyle -eq $xlLineStyleContinuous -and $rightDisplayStyle -eq $xlLineStyleContinuous) {
                if ($leftDirect -eq $xlLineStyleContinuous -and $rightDirect -eq $xlLineStyleContinuous) {
                    'DIRECT'
                }
                else {
                    'STYLE-ONLY'
                }
            }
            else {
                'MISSING'
            }

            Write-Host ('  row {0,3} h={1,5} direct={2}/{3}/{4} display={5}({6})/{7}({8}) fill={9}/{10} merged={11,-5} => {12}' -f `
                    $row,
                    (Format-Number $Sheet.Rows.Item($row).RowHeight),
                    $leftDirect, $rightDirect, $pairDirect,
                    $leftDisplayStyle, (Format-OleColour $leftDisplay.Color),
                    $rightDisplayStyle, (Format-OleColour $rightDisplay.Color),
                    $leftFill, $rightFill,
                    $merged,
                    $verdict)

            if ($verdict -ne 'DIRECT') {
                $findings += ('row {0,3} {1,-22} {2} (direct={3}/{4}/{5}, display={6}/{7}, fill={8}/{9})' -f `
                        $row, $name, $verdict, $leftDirect, $rightDirect, $pairDirect,
                        $leftDisplayStyle, $rightDisplayStyle, $leftFill, $rightFill)
            }
        }

        Write-Host ''
    }

    Write-Host '=== 3. Summary: boundaries the host would NOT paint as a direct rule ==='
    if ($findings.Count -eq 0) {
        Write-Host 'None. Every visible boundary reads DIRECT on every row, so the model is healthy and'
        Write-Host 'a visible gap is a painting problem rather than a formatting problem.'
    }
    else {
        foreach ($finding in $findings) {
            Write-Host ('  ' + $finding)
        }

        Write-Host ''
        Write-Host 'Reading the findings:'
        Write-Host '  STYLE-ONLY on body rows only  -> the table style supplies the rule and the appended'
        Write-Host '                                   rows carry no direct border (insert format origin).'
        Write-Host '  MISSING on body rows only     -> those rows carry a direct no-border format, which'
        Write-Host '                                   suppresses the style. Same insert path, stronger form.'
        Write-Host '  DIRECT but colour == fill     -> the rule is painted in the fill''s own colour: a'
        Write-Host '                                   colour defect, invisible on screen, green in LineStyle.'
        Write-Host '  MERGED                        -> the two cells are one range; no boundary can exist.'
    }

    Write-Host ''
}

$excel = $null
$workbook = $null
$ownsExcel = $false
try {
    if ($Path) {
        if (-not (Test-Path $Path)) {
            Write-Host "No such workbook: $Path" -ForegroundColor Red
            return
        }

        # Excel resolves a RELATIVE path against its own working directory, not PowerShell's,
        # and reports the classic "couldn't find ... moved, renamed or deleted" dialog. The
        # path is made absolute before the host sees it.
        $fullPath = (Resolve-Path -Path $Path).ProviderPath

        $excel = New-Object -ComObject Excel.Application
        $ownsExcel = $true
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($fullPath, 0, $true)
        Write-Host "Opened read-only in a private instance: $fullPath"
    }
    else {
        # Marshal.GetActiveObject does not exist on .NET (PowerShell 7), which is why this
        # uses the VisualBasic ROT lookup instead of the classic .NET Framework call.
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        $excel = [Microsoft.VisualBasic.Interaction]::GetObject($null, 'Excel.Application')

        if ($excel.Workbooks.Count -eq 0) {
            Write-Host 'Attached to a running Excel that has no workbook open.' -ForegroundColor Yellow
            Write-Host 'Pass -Path <workbook>, or open the Gantt workbook first.'
            return
        }

        Write-Host "Attached to the running Excel: $($excel.Version)"
        $workbook = $excel.ActiveWorkbook
    }

    Write-Host "Excel version: $($excel.Version)"
    if ($null -eq $workbook) {
        Write-Host 'There is no active workbook.' -ForegroundColor Red
        return
    }

    Write-Host "Workbook:      $($workbook.Name)"

    $sheet = if ($SheetName) { $workbook.Worksheets.Item($SheetName) } else { Find-GanttSheet -Workbook $workbook }
    if ($null -eq $sheet) {
        Write-Host "No worksheet carries the table '$dataTableName'." -ForegroundColor Red
        return
    }

    $table = Find-GanttTable -Sheet $sheet
    Write-Host "Gantt sheet:   $($sheet.Name)"
    Write-Host ''

    Write-TableFacts -Table $table -Sheet $sheet
    Write-BoundaryMatrix -Table $table -Sheet $sheet

    Write-Host 'PROBE COMPLETE'
}
finally {
    # An attached instance is never quit and its workbook is never closed: the owner may
    # have unsaved work in it. A private instance is always closed and quit, and the
    # workbook is never saved.
    if ($ownsExcel -and $null -ne $excel) {
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    elseif ($null -ne $excel) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
