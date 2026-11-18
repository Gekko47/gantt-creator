#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe (R5.18n D4 phase C0): does finalising BEFORE Quit() change
    whether Excel terminates?

.DESCRIPTION
    OfficeFixture calls Quit() and only then runs
    GC.Collect/WaitForPendingFinalizers/GC.Collect. An RCW a body left unreleased
    but UNREACHABLE is finalisable, and its COM reference is still held at the
    instant Quit() runs unless a finaliser pass has already completed -- which
    would make the forced-kill outcome a race rather than a per-adapter property.

    This probe models both orders side by side, ten repeats each, with the
    leaking-body shape: a heavy workload (28 columns, whole-column ranges, fonts,
    hidden columns, a defined name) built inside a function whose scope dies before
    teardown, with NO explicit releases. Both orders keep the post-Quit GC pass, so
    the only difference is the one before Quit.

        Order A (current fixture): Quit -> GC -> poll
        Order B (proposed):        GC -> Quit -> GC -> poll

.RESULTS
    Measured 2026-11-18 on Excel 16.0.20527.20196:

        order A (Quit -> GC): 10/10 exited (two runs at ~2.3-2.6 s, eight at ~0.26 s)
        order B (GC -> Quit): 10/10 exited (~2.3 s each)

    So the ordering hypothesis is NOT supported: with deliberately leaked,
    unreachable proxies the process leaves either way. A companion measurement
    (`dotnet test --filter` on one leaking body) still read **1 forced kill at
    10,098 ms**, which rules out accumulation across the run and shows the flap is
    deterministic per body INSIDE the xUnit host and does not reproduce under
    PowerShell COM at all. The discriminator therefore has to be collected inside
    the harness (Quit outcome, body duration, whether the workbook was initialised),
    not from a standalone probe. Recorded in docs/work-items/R5.18n-deferred-close-out.md.

.NOTES
    Disposable spike, not part of the gate. Run with:

        pwsh ./scripts/probe-quit-ordering.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:pollSeconds = 10

# The leaking-body shape: every proxy acquired here becomes unreachable when the
# function returns, and none of them is released. That is what a test body hands
# to teardown.
function Initialize-HeavyWorkload {
    param([Parameter(Mandatory)] $Excel)

    $book = $Excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)

    foreach ($index in 1..28) {
        $sheet.Cells.Item(1, $index).Value2 = "Col$index"
    }

    foreach ($row in 2..6) {
        foreach ($index in 1..28) {
            $sheet.Cells.Item($row, $index).Value2 = "$row-$index"
        }
    }

    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:AB6'), $null, 1)
    foreach ($index in 1..28) {
        $column = $table.ListColumns.Item($index)
        $range = $column.Range
        $entire = $range.EntireColumn
        $entire.ColumnWidth = 9
        $entire.Font.Size = 9
        $entire.HorizontalAlignment = -4108
        if ($index -le 10) {
            $entire.Hidden = $true
        }
    }

    $null = $book.Names.Add('GanttPlotAnchor', '=Sheet1!$A$1')
    $book.Close($false)

    # Deliberately NOTHING here releases $book, $sheet, $table, $column, $range or
    # $entire: the leak the probe is testing for.
}

function Invoke-OrderedCase {
    param(
        [Parameter(Mandatory)][string] $Order,
        [Parameter(Mandatory)][int] $Repeat
    )

    $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = $null
    $exited = $false
    $elapsedMs = -1

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        $excel.Visible = $false
        Initialize-HeavyWorkload -Excel $excel

        if ($Order -eq 'B') {
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
        }

        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()

        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $before -notcontains $_.Id })
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while ($remaining.Count -gt 0 -and $watch.Elapsed.TotalSeconds -lt $script:pollSeconds) {
            Start-Sleep -Milliseconds 250
            $remaining = @($remaining | Where-Object { Get-Process -Id $_.Id -ErrorAction SilentlyContinue })
        }

        $elapsedMs = [int]$watch.ElapsedMilliseconds
        if ($remaining.Count -gt 0) {
            foreach ($process in $remaining) {
                & taskkill /PID $process.Id /T /F 2>$null | Out-Null
            }
        } else {
            $exited = $true
        }
    } catch {
        Write-Host ("  order {0} repeat {1,-2} ERROR: {2}" -f $Order, $Repeat, $_.Exception.Message)
        return $false
    } finally {
        if ($null -ne $excel) {
            try {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            } catch {
                Write-Verbose "release refused: $($_.Exception.Message)"
            }
        }
    }

    $verdict = if ($exited) { 'EXITED' } else { 'killed at the bound' }
    Write-Host ("order {0} repeat {1,-2} {2,-22} {3} ms" -f $Order, $Repeat, $verdict, $elapsedMs)
    return $exited
}

$excelVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).VersionToReport
Write-Host "Office $excelVersion; 10 repeats per order, Quit polled for $script:pollSeconds s"
Write-Host 'Order A = Quit then GC (current fixture). Order B = GC then Quit (proposed).'
Write-Host ''

$exitsA = 0
foreach ($repeat in 1..10) {
    if (Invoke-OrderedCase -Order 'A' -Repeat $repeat) { $exitsA++ }
}

Write-Host ''
$exitsB = 0
foreach ($repeat in 1..10) {
    if (Invoke-OrderedCase -Order 'B' -Repeat $repeat) { $exitsB++ }
}

Write-Host ''
Write-Host ("SUMMARY  order A (Quit->GC): {0}/10 exited    order B (GC->Quit): {1}/10 exited" -f $exitsA, $exitsB)
Write-Host 'Done. Ordinary probe output; this script is disposable.'
