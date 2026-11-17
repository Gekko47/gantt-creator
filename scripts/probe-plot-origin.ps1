#Requires -Version 7.0
<#
.SYNOPSIS
    Read-only probe: where does the LIVE plot actually start, and where is the
    plot's left edge line anchored?

.DESCRIPTION
    Owner report (2026-11-15): "the plot is still plotting from column G even though
    the anchor is in S" and "the plot's left edge is not visible in the normal view".
    Two authorities can decide the plot's left edge, and they only disagree once any
    engine column is unhidden:

      * PanelCellGrid.PanelRightPt returns the MEASURED left edge of the plot anchor
        column (column S, GanttSheetLayout.PlotAnchorColumnIndex), and falls back to
        OriginLeftPt + TotalWidthPt -- the summed widths of the columns the add-in
        measured, which are the five VISIBLE table columns.

    While the engine columns are hidden the two are the same number (a hidden column
    has zero width), so the normal view cannot tell them apart. That is exactly why
    the owner's "all columns unhidden" screenshot is the diagnostic one.

    This probe prints, for the Gantt sheet in a workbook:

      1. every worksheet column's Left / Width / Hidden for A..V,
      2. the table's header and body ranges,
      3. the three candidate plot origins, computed the way the add-in computes
         them, and
      4. every chart shape's Left, Top, Width, Height, TopLeftCell and Placement --
         the TopLeftCell is what proves whether the plot's left edge is anchored
         inside a hidden column (Excel does not paint a shape anchored there).

    It never writes to the workbook. With -Path it opens a private, invisible Excel
    instance and closes it again; without -Path it attaches to the Excel that is
    already running and leaves it running.

.PARAMETER Path
    Full path to the workbook to inspect. Opened read-only in a new, invisible
    Excel instance that is closed before the script returns.

.EXAMPLE
    pwsh ./scripts/probe-plot-origin.ps1 -Path 'C:\Users\rossj\Documents\gantt.xlsx'

.EXAMPLE
    # The Gantt workbook is already open: attach to that Excel and read it.
    pwsh ./scripts/probe-plot-origin.ps1

.NOTES
    Disposable spike, in the style of probe-frame-anchor.ps1. Results are printed as
    they are measured, because values built from COM proxies do not survive a
    PowerShell function boundary reliably.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string] $Path
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$dataTableName = 'tblGanttData'
$anchorColumnIndex = 19          # GanttSheetLayout.PlotAnchorColumnIndex (S)
$firstTableColumnIndex = 2       # GanttSheetLayout.TableOriginColumnIndex (B)
$lastTableColumnIndex = 18       # 17 schema columns from B: B..R
$visibleColumnCount = 5          # Type, Description, Start, Finish, Duration
$anchorNudgePt = 0.25            # GanttSheetLayout.LiveFrameAnchorNudgePt

function Format-Num {
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

function Write-Columns {
    param([Parameter(Mandatory)] $Sheet)

    Write-Host '=== 1. Worksheet columns (Left is measured from the sheet edge, in points) ==='
    Write-Host ('{0,-4} {1,-5} {2,10} {3,10} {4,8} {5,12}' -f 'Idx', 'Name', 'Left', 'Width', 'Hidden', 'ColumnWidth')
    for ($index = 1; $index -le 22; $index++) {
        $column = $Sheet.Columns.Item($index)
        $letter = $column.Address($false, $false).Split(':')[0]
        Write-Host ('{0,-4} {1,-5} {2,10} {3,10} {4,8} {5,12}' -f `
                $index,
            $letter,
            (Format-Num $column.Left),
            (Format-Num $column.Width),
            $column.Hidden,
            (Format-Num $column.ColumnWidth))
    }

    Write-Host ''
}

function Write-Origins {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] $Table
    )

    Write-Host '=== 2. Table ranges ==='
    $header = $Table.HeaderRowRange
    Write-Host "Header row:    $($header.Address($false, $false))  Left=$(Format-Num $header.Left)  Width=$(Format-Num $header.Width)  Top=$(Format-Num $header.Top)  Height=$(Format-Num $header.Height)"
    $body = $Table.DataBodyRange
    if ($null -eq $body) {
        Write-Host 'Body range:    <none - the table is empty (ADR-0050 empty-chart state)>' -ForegroundColor Yellow
        $originLeft = [double]$header.Left
    }
    else {
        Write-Host "Body range:    $($body.Address($false, $false))  Left=$(Format-Num $body.Left)  Width=$(Format-Num $body.Width)  Top=$(Format-Num $body.Top)  Height=$(Format-Num $body.Height)"
        $originLeft = [double]$body.Left
    }

    Write-Host ''
    Write-Host '=== 3. The candidate plot origins, in points ==='

    $visibleSum = 0.0
    $tableSum = 0.0
    for ($index = $firstTableColumnIndex; $index -le $lastTableColumnIndex; $index++) {
        $width = [double]$Sheet.Columns.Item($index).Width
        $tableSum += $width
        if ($index -lt ($firstTableColumnIndex + $visibleColumnCount)) {
            $visibleSum += $width
        }
    }

    $anchorLeft = [double]$Sheet.Columns.Item($anchorColumnIndex).Left

    Write-Host "Table origin (measured Left of the header/body):  $(Format-Num $originLeft)"
    Write-Host "Fallback: origin + the 5 VISIBLE columns:         $(Format-Num ($originLeft + $visibleSum))"
    Write-Host "Reference: origin + all 17 table columns:          $(Format-Num ($originLeft + $tableSum))"
    Write-Host "ANCHOR: measured Left of column S:                 $(Format-Num $anchorLeft)"
    Write-Host "ANCHOR + nudge (what the add-in renders):          $(Format-Num ($anchorLeft + $anchorNudgePt))"
    Write-Host ''
    Write-Host ("Does the fallback equal the anchor column?         {0}" -f ([Math]::Abs(($originLeft + $visibleSum) - $anchorLeft) -lt 0.01))
    Write-Host ''
}

function Write-Shapes {
    param([Parameter(Mandatory)] $Sheet)

    Write-Host '=== 4. Chart shapes (sorted by Left, then Top) ==='
    $shapes = @()
    $shapeCount = $Sheet.Shapes.Count
    for ($shapeIndex = 1; $shapeIndex -le $shapeCount; $shapeIndex++) {
        $shape = $Sheet.Shapes.Item($shapeIndex)
        if (-not $shape.Name.StartsWith('chart:', [StringComparison]::Ordinal)) {
            continue
        }

        $shapes += [pscustomobject]@{
            Name    = $shape.Name
            Left    = [double]$shape.Left
            Top     = [double]$shape.Top
            Width   = [double]$shape.Width
            Height  = [double]$shape.Height
            Anchor  = $shape.TopLeftCell.Address($false, $false)
            Place   = $shape.Placement
            Visible = $shape.Visible
        }
    }

    Write-Host ("chart:* shape count: {0}" -f $shapes.Count)
    Write-Host ('{0,-34} {1,10} {2,10} {3,9} {4,9} {5,-8} {6,10} {7,8}' -f 'Name', 'Left', 'Top', 'Width', 'Height', 'Anchor', 'Placement', 'Visible')
    foreach ($shape in ($shapes | Sort-Object Left, Top | Select-Object -First 40)) {
        Write-Host ('{0,-34} {1,10} {2,10} {3,9} {4,9} {5,-8} {6,10} {7,8}' -f `
                $shape.Name,
            (Format-Num $shape.Left),
            (Format-Num $shape.Top),
            (Format-Num $shape.Width),
            (Format-Num $shape.Height),
            $shape.Anchor,
            $shape.Place,
            $shape.Visible)
    }

    if ($shapes.Count -gt 40) {
        Write-Host ("... {0} further chart shapes not shown" -f ($shapes.Count - 40))
    }

    Write-Host ''
    Write-Host '=== 5. The plot''s own left edge ==='
    $edges = $shapes | Where-Object { $_.Name -like 'chart:grid:edge:*' -or $_.Name -like 'chart:frame:left' -or $_.Name -like 'chart:plot:*' }
    if ($null -eq $edges) {
        Write-Host 'No chart:grid:edge / chart:frame:left / chart:plot shape is present.' -ForegroundColor Yellow
    }
    else {
        foreach ($shape in ($edges | Sort-Object Left, Top)) {
            Write-Host ('{0,-34} Left={1,9} Top={2,9} W={3,8} H={4,8} Anchor={5}' -f `
                    $shape.Name, (Format-Num $shape.Left), (Format-Num $shape.Top), (Format-Num $shape.Width), (Format-Num $shape.Height), $shape.Anchor)
        }
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

        # Excel resolves a RELATIVE path against its own working directory, not
        # PowerShell's, and reports the classic "couldn't find ... moved, renamed or
        # deleted" dialog. The path is made absolute before the host sees it.
        $fullPath = (Resolve-Path -Path $Path).ProviderPath

        $excel = New-Object -ComObject Excel.Application
        $ownsExcel = $true
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($fullPath, 0, $true)
        Write-Host "Opened read-only in a private instance: $fullPath"
    }
    else {
        # Marshal.GetActiveObject does not exist on .NET (PowerShell 7), which is why
        # this uses the VisualBasic ROT lookup instead of the classic .NET Framework call.
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
        $excel = [Microsoft.VisualBasic.Interaction]::GetObject($null, 'Excel.Application')
        $workbook = $excel.ActiveWorkbook

        if ($excel.Workbooks.Count -eq 0) {
            Write-Host 'Attached to a running Excel that has no workbook open.' -ForegroundColor Yellow
            Write-Host 'Pass -Path <workbook>, or open the Gantt workbook first.'
            return
        }

        Write-Host "Attached to the running Excel: $($excel.Version)"
    }

    Write-Host "Excel version: $($excel.Version)"

    if ($null -eq $workbook) {
        Write-Host 'There is no active workbook.' -ForegroundColor Red
        return
    }

    Write-Host "Workbook:      $($workbook.Name)"

    $sheet = Find-GanttSheet -Workbook $workbook
    if ($null -eq $sheet) {
        Write-Host "No worksheet carries the table '$dataTableName'." -ForegroundColor Red
        Write-Host 'The Gantt sheet must exist in this workbook.'
        return
    }

    $table = Find-GanttTable -Sheet $sheet
    Write-Host "Gantt sheet:   $($sheet.Name)"
    Write-Host ''

    Write-Columns -Sheet $sheet
    Write-Origins -Sheet $sheet -Table $table
    Write-Shapes -Sheet $sheet

    Write-Host 'PROBE COMPLETE'
}
finally {
    # An attached instance is never quit and its workbook is never closed: the owner
    # may have unsaved work in it. A private instance is always closed and quit, and
    # the workbook is never saved.
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
