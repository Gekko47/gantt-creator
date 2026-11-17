#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable live-host probe: can Excel hide every user-visible table column of
    tblGanttData, and what does it report afterwards?

.DESCRIPTION
    R5.22 makes "no visible table columns" a supported state. Three facts cannot be
    settled from source code:

      1. does the host accept EntireColumn.Hidden = $true for all six user columns?
      2. with every table column hidden, does Range.Left still report the table origin,
         and does the plot anchor column's Left equal that origin?
      3. where does the active cell go when the column containing it is hidden?

    The probe opens a THROWAWAY workbook in a private, invisible Excel instance,
    records the baseline, hides the six user columns by header name, reads the live
    values back, restores the original hidden flags, and closes WITHOUT SAVING.

    It is a spike, not a gate: it prints evidence and exits non-zero only when the host
    refuses the all-hidden state or a measurement the add-in depends on reads null.

.PARAMETER Path
    Full path to a throwaway workbook that contains tblGanttData. The workbook is
    never saved. Omit it when -Build is used.

.PARAMETER Build
    Creates a synthetic schema-26 workbook in the temp folder (six user columns plus
    twelve engine columns, two body rows), probes it, then deletes it. This is the
    reproducible evidence path when no current workbook is available.

.PARAMETER KeepHidden
    Restores nothing. Use only when the caller wants to inspect the host afterwards;
    the workbook is still closed without saving.

.EXAMPLE
    pwsh ./scripts/probe-hide-all-columns.ps1 -Build

.EXAMPLE
    pwsh ./scripts/probe-hide-all-columns.ps1 -Path ./scripts/_artifacts/plot-origin-fixture.xlsx

.NOTES
    Disposable spike in the style of probe-plot-origin.ps1. The six display names are
    the R5.21 visible set; an engine column is never touched.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $Path,

    [switch] $Build,

    [switch] $KeepHidden
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dataTableName = 'tblGanttData'
$anchorColumnIndex = 0           # derived from the live table: origin + column count
$userColumns = @(
    'Type'
    'Activity ID'
    'Description'
    'Start'
    'Finish'
    'Duration'
)
$schemaColumns = @(
    'Type'
    'Activity ID'
    'Description'
    'Start'
    'Finish'
    'Duration'
    'Id'
    'LaneId'
    'StackIndex'
    'ParentId'
    'SiblingOrder'
    'StyleKey'
    'LabelPosition'
    'DateLabelPosition'
    'FillColour'
    'StrokeColour'
    'Visible'
    'SortOrder'
)
$tolerance = 0.01

function Format-Num {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return 'null'
    }

    return [Math]::Round([double]$Value, 3).ToString('0.###')
}

function Get-Num {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    return [double]$Value
}

function Find-GanttSheet {
    param([Parameter(Mandatory)] $Workbook)

    $count = $Workbook.Worksheets.Count
    for ($index = 1; $index -le $count; $index++) {
        $sheet = $Workbook.Worksheets.Item($index)
        $tables = $sheet.ListObjects
        $tableCount = $tables.Count
        for ($tableIndex = 1; $tableIndex -le $tableCount; $tableIndex++) {
            if ($tables.Item($tableIndex).Name -eq $dataTableName) {
                return $sheet
            }
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

function Resolve-UserColumns {
    param([Parameter(Mandatory)] $Table)

    $resolved = [ordered]@{}
    $columns = $Table.ListColumns
    for ($index = 1; $index -le $columns.Count; $index++) {
        $column = $columns.Item($index)
        if ($userColumns -contains $column.Name) {
            $resolved[$column.Name] = $column
        }
    }

    return $resolved
}

function Write-Baseline {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] $Table,
        [Parameter(Mandatory)] $Resolved
    )

    $header = $Table.HeaderRowRange
    $body = $Table.DataBodyRange
    $anchor = $Sheet.Columns.Item($anchorColumnIndex)

    Write-Host '=== Baseline ==='
    Write-Host "Header range:  $($header.Address($false, $false))"
    Write-Host "  Left=$(Format-Num $header.Left)  Width=$(Format-Num $header.Width)"
    if ($null -eq $body) {
        Write-Host 'Body range:    <none - empty table>'
    }
    else {
        Write-Host "Body range:    $($body.Address($false, $false))  Rows=$($body.Rows.Count)"
        Write-Host "  Left=$(Format-Num $body.Left)  Width=$(Format-Num $body.Width)"
    }

    Write-Host "Anchor column ${anchorColumnIndex}: Left=$(Format-Num $anchor.Left)  Width=$(Format-Num $anchor.Width)"
    Write-Host "Active cell:   $($excel.ActiveCell.Address($false, $false))"
    Write-Host "Active sheet:  $($excel.ActiveSheet.Name)"
    Write-Host "Freeze panes:  $($excel.ActiveWindow.FreezePanes)  SplitColumn=$($excel.ActiveWindow.SplitColumn)  SplitRow=$($excel.ActiveWindow.SplitRow)"
    Write-Host ''

    Write-Host 'User columns:'
    foreach ($name in $userColumns) {
        if (-not $Resolved.Contains($name)) {
            Write-Host ("  {0,-12} MISSING" -f $name) -ForegroundColor Red
            continue
        }

        $column = $Resolved[$name]
        $entire = $column.Range.EntireColumn
        Write-Host ("  {0,-12} Index={1,2} Hidden={2,-5} Left={3,8} Width={4,8}" -f `
                $name,
            $column.Index,
            $entire.Hidden,
            (Format-Num $entire.Left),
            (Format-Num $entire.Width))
    }

    Write-Host ''
}

function New-ProbeWorkbook {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] $Workbook
    )

    if (-not $PSCmdlet.ShouldProcess('synthetic workbook', 'build the schema-26 probe table')) {
        return
    }

    $sheet = $Workbook.Worksheets.Item(1)
    $sheet.Name = 'Gantt Data'

    for ($index = 0; $index -lt $schemaColumns.Count; $index++) {
        $sheet.Cells.Item(1, 2 + $index).Value2 = $schemaColumns[$index]
    }

    # All-string cells: the probe measures column visibility and geometry, not date
    # typing, and a homogeneous string[] avoids PowerShell's array type inference.
    $rows = @(
        @('Activity', 'A1', 'Probe row 1', '2026-01-01', '2026-01-03', '2', '1', '1', '1', '', '1', '', '', '', '', '', 'TRUE', '1'),
        @('Activity', 'A2', 'Probe row 2', '2026-01-02', '2026-01-05', '3', '2', '1', '1', '1', '2', '', '', '', '', '', 'TRUE', '2')
    )

    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        for ($columnIndex = 0; $columnIndex -lt $schemaColumns.Count; $columnIndex++) {
            $sheet.Cells.Item(2 + $rowIndex, 2 + $columnIndex).Value2 = $rows[$rowIndex][$columnIndex]
        }
    }

    $range = $sheet.Range(
        $sheet.Cells.Item(1, 2),
        $sheet.Cells.Item(1 + $rows.Count, 1 + $schemaColumns.Count))
    $table = $sheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $dataTableName

    # Column A is the leading gutter; the six user columns are readable, the twelve
    # engine columns are narrow and hidden, exactly as Initialise presents them.
    $sheet.Columns.Item(1).ColumnWidth = 2.0
    for ($index = 0; $index -lt $userColumns.Count; $index++) {
        $sheet.Columns.Item(2 + $index).ColumnWidth = 8.43
    }
    for ($index = $userColumns.Count; $index -lt $schemaColumns.Count; $index++) {
        $sheet.Columns.Item(2 + $index).ColumnWidth = 2.0
        $sheet.Columns.Item(2 + $index).Hidden = $true
    }

    $sheet.Activate()
}

$excel = $null
$workbook = $null
$ownsExcel = $false
$originalHidden = @{}
$pass = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $ownsExcel = $true
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    if ($Build) {
        $workbook = $excel.Workbooks.Add()
        New-ProbeWorkbook -Excel $excel -Workbook $workbook
        Write-Host 'Built a synthetic schema-26 workbook in memory (never saved).'
    }
    else {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            throw 'Pass -Path <workbook> or -Build.'
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            Write-Host "No such workbook: $Path" -ForegroundColor Red
            exit 1
        }

        $fullPath = (Resolve-Path -LiteralPath $Path).ProviderPath
        $workbook = $excel.Workbooks.Open($fullPath, 0, $false)
        Write-Host "Opened for probing (never saved): $fullPath"
    }

    Write-Host "Excel version: $($excel.Version)"
    Write-Host "Workbook: $($workbook.Name)"

    $sheet = Find-GanttSheet -Workbook $workbook
    if ($null -eq $sheet) {
        Write-Host "No worksheet carries the table '$dataTableName'." -ForegroundColor Red
        exit 1
    }

    $table = Find-GanttTable -Sheet $sheet
    $anchorColumnIndex = [int]$table.Range.Column + [int]$table.ListColumns.Count
    $sheet.Activate()
    Write-Host "Gantt sheet: $($sheet.Name)"
    Write-Host "Anchor column: $anchorColumnIndex (derived from the live table)"
    Write-Host ''

    $resolved = Resolve-UserColumns -Table $table
    Write-Baseline -Sheet $sheet -Table $table -Resolved $resolved

    $missing = @($userColumns | Where-Object { -not $resolved.Contains($_) })
    if ($missing.Count -gt 0) {
        Write-Host ("Cannot probe: the table is missing {0}." -f ($missing -join ', ')) -ForegroundColor Red
        exit 1
    }

    # Remember every user column's state so the workbook can be restored before closing.
    foreach ($name in $userColumns) {
        $entire = $resolved[$name].Range.EntireColumn
        $originalHidden[$name] = [bool]$entire.Hidden
    }

    # Put the active cell inside the first user column, so hiding it exercises the
    # host's selection move rather than a cell that is already elsewhere.
    # ListColumn.Index is TABLE-relative; the worksheet column comes from Range.Column.
    $headerRow = [int]$table.HeaderRowRange.Row
    $firstColumnIndex = [int]$resolved['Type'].Range.Column
    $null = $sheet.Cells.Item($headerRow + 1, $firstColumnIndex).Select()
    Write-Host "Active cell before hide: $($excel.ActiveCell.Address($false, $false))"
    Write-Host ''

    Write-Host '=== Hiding every user column ==='
    $hideFailure = $null
    foreach ($name in $userColumns) {
        $entire = $resolved[$name].Range.EntireColumn
        try {
            $entire.Hidden = $true
            Write-Host ("  {0,-12} Hidden written" -f $name)
        }
        catch {
            $hideFailure = $_
            Write-Host ("  {0,-12} REFUSED: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            break
        }
    }
    Write-Host ''

    if ($null -ne $hideFailure) {
        Write-Host 'PROBE RESULT: FAIL - the host refused to hide a user column.' -ForegroundColor Red
        Write-Host ("HRESULT: {0}" -f $hideFailure.Exception.HResult)
        exit 1
    }

    $headerAfter = $table.HeaderRowRange
    $bodyAfter = $table.DataBodyRange
    $anchorAfter = $sheet.Columns.Item($anchorColumnIndex)

    $hiddenRows = @()
    foreach ($name in $userColumns) {
        $entire = $resolved[$name].Range.EntireColumn
        $hiddenRows += [pscustomobject]@{
            Name    = $name
            Hidden  = $entire.Hidden
            Left    = Get-Num $entire.Left
            Width   = Get-Num $entire.Width
        }
    }

    $headerLeft = Get-Num $headerAfter.Left
    $headerWidth = Get-Num $headerAfter.Width
    $bodyLeft = if ($null -eq $bodyAfter) { $null } else { Get-Num $bodyAfter.Left }
    $bodyWidth = if ($null -eq $bodyAfter) { $null } else { Get-Num $bodyAfter.Width }
    $bodyRows = if ($null -eq $bodyAfter) { 0 } else { [int]$bodyAfter.Rows.Count }
    $anchorLeft = Get-Num $anchorAfter.Left
    $anchorWidth = Get-Num $anchorAfter.Width
    $originLeft = if ($null -ne $bodyLeft) { $bodyLeft } else { $headerLeft }
    $activeCell = [string]$excel.ActiveCell.Address($false, $false)
    $activeSheet = [string]$excel.ActiveSheet.Name

    Write-Host '=== After hiding every user column ==='
    Write-Host ('{0,-12} {1,-7} {2,10} {3,10}' -f 'Column', 'Hidden', 'Left', 'Width')
    foreach ($row in $hiddenRows) {
        Write-Host ('{0,-12} {1,-7} {2,10} {3,10}' -f `
                $row.Name, $row.Hidden, (Format-Num $row.Left), (Format-Num $row.Width))
    }

    Write-Host ''
    Write-Host "Header: Left=$(Format-Num $headerLeft)  Width=$(Format-Num $headerWidth)"
    Write-Host "Body:   Left=$(Format-Num $bodyLeft)  Width=$(Format-Num $bodyWidth)  Rows=$bodyRows  Present=$($null -ne $bodyAfter)"
    Write-Host "Anchor: Left=$(Format-Num $anchorLeft)  Width=$(Format-Num $anchorWidth)"
    Write-Host "Active cell: $activeCell"
    Write-Host "Active sheet: $activeSheet"
    Write-Host "Freeze panes: $($excel.ActiveWindow.FreezePanes)  SplitColumn=$($excel.ActiveWindow.SplitColumn)  SplitRow=$($excel.ActiveWindow.SplitRow)"
    Write-Host ''

    $allHidden = @($hiddenRows | Where-Object { $_.Hidden -ne $true }).Count -eq 0
    $widthsZero = @($hiddenRows | Where-Object { $null -eq $_.Width -or [Math]::Abs($_.Width) -gt $tolerance }).Count -eq 0
    $originReadable = $null -ne $originLeft
    $originEqualsAnchor = $originReadable -and $null -ne $anchorLeft -and
        [Math]::Abs($originLeft - $anchorLeft) -le $tolerance
    $anchorVisible = $null -ne $anchorWidth -and $anchorWidth -gt 0
    $bodyReadable = $null -ne $bodyAfter
    $activeCellReadable = -not [string]::IsNullOrWhiteSpace($activeCell)

    Write-Host '=== Verdict ==='
    Write-Host ("All six columns report Hidden=true:        {0}" -f $allHidden)
    Write-Host ("All six hidden widths report zero:         {0}" -f $widthsZero)
    Write-Host ("Header/body Left is readable:              {0} (origin={1})" -f $originReadable, (Format-Num $originLeft))
    Write-Host ("Origin equals the anchor column's Left:    {0}" -f $originEqualsAnchor)
    Write-Host ("Anchor column is still visible:            {0} (width={1})" -f $anchorVisible, (Format-Num $anchorWidth))
    Write-Host ("DataBodyRange still exists:                {0} (rows={1})" -f $bodyReadable, $bodyRows)
    Write-Host ("Active cell is readable:                   {0} ({1})" -f $activeCellReadable, $activeCell)
    Write-Host ''

    $pass = $allHidden -and $widthsZero -and $originReadable -and $originEqualsAnchor -and $anchorVisible -and $bodyReadable -and $activeCellReadable
    if ($pass) {
        Write-Host 'PROBE RESULT: PASS - the host supports the all-hidden state and every measurement the add-in needs is readable.' -ForegroundColor Green
    }
    else {
        Write-Host 'PROBE RESULT: FAIL - one or more facts the add-in depends on did not hold. Record this and stop before coding.' -ForegroundColor Red
    }
}
catch {
    Write-Host "PROBE ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    $pass = $false
}
finally {
    if (-not $KeepHidden -and $null -ne $workbook -and $originalHidden.Count -gt 0) {
        try {
            foreach ($name in $userColumns) {
                if (-not $originalHidden.ContainsKey($name)) {
                    continue
                }

                $entire = $resolved[$name].Range.EntireColumn
                $entire.Hidden = [bool]$originalHidden[$name]
            }
        }
        catch {
            Write-Host "Could not restore the original visibility: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    if ($ownsExcel -and $null -ne $excel) {
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }

        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not $pass) {
    exit 1
}

exit 0
