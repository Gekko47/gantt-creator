#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe (R5.18n D4): does any single Excel operation make Quit()
    fail, or is the forced-kill flap independent of what the body did?

.DESCRIPTION
    The Office gate records exactly one forced kill for most workbook bodies, and
    the bodies that read clean are the ones that never run
    ExcelWorkbookInitialiser.Initialise. Before blaming an adapter's unreleased
    proxies, this probe tests the alternative explanation: that the process
    sometimes refuses to terminate for a reason that is neither a proxy nor a
    specific operation.

    Each case starts a FRESH Excel, performs ONE operation, closes the workbook,
    releases the proxies exactly as OfficeFixture does, runs the same GC pass, then
    polls Quit for 10 s. A case still running after 10 s is killed and reported.

    A SECOND section repeats the table case five times and forces calculation
    before the close, because the first run's two "stuck" cases were both table
    cases and only an out-of-band repeat can tell a cause from a coincidence.

.RESULTS
    Measured 2026-11-18 on Excel 16.0.20527.20196, two consecutive runs:

    Run 1: control, row-height, Names.Add, table+Names, table+validation,
    table+shape, second-sheet and protect-release all EXITED; `ListObjects.Add`
    and `table + hidden column` were STILL RUNNING at 10 s and were killed.

    Run 2 (identical script): **every case EXITED**, including both that had hung,
    and all five table repeats and both calculate variants exited. Timings split
    into two bands, about 260 ms and about 2.8 s, with no correlation to the
    operation.

    Conclusion recorded in docs/work-items/R5.18n-deferred-close-out.md: the flap
    is non-deterministic at the COM level and is NOT attributable to a single
    operation, so "release the initialiser's proxies" was not evidence-led and was
    withdrawn. The probe did show that a table, a defined name, a validation, a
    shape, a second sheet and a protect/release cycle all terminate cleanly on
    their own.

.NOTES
    Disposable spike, not part of the gate. Run with:

        pwsh ./scripts/probe-quit-cases.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Invoke-QuitCase {
    param(
        [Parameter(Mandatory)][string] $Name,
        [Parameter(Mandatory)][scriptblock] $Setup
    )

    $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $excel = $null
    $book = $null
    $exited = $false
    $elapsedMs = -1

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.DisplayAlerts = $false
        $excel.Visible = $false
        $book = $excel.Workbooks.Add()
        & $Setup $excel $book

        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
        $book = $null

        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        $excel = $null

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()

        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $before -notcontains $_.Id })
        $watch = [Diagnostics.Stopwatch]::StartNew()
        while ($remaining.Count -gt 0 -and $watch.Elapsed.TotalSeconds -lt 10) {
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
        Write-Host ("  {0,-46} ERROR: {1}" -f $Name, $_.Exception.Message)
        return
    } finally {
        if ($null -ne $excel) {
            try {
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            } catch {
                # A refused release cannot change the probe's measurement; the
                # process is either gone or already queued for the kill below.
                Write-Verbose "release refused: $($_.Exception.Message)"
            }
        }
    }

    $verdict = if ($exited) { 'EXITED' } else { 'STILL RUNNING after 10 s (killed)' }
    Write-Host ("{0,-46} {1}  ({2} ms)" -f $Name, $verdict, $elapsedMs)
}

$excelVersion = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).VersionToReport
Write-Host "Office $excelVersion; each case is a fresh Excel, Quit polled for 10 s"
Write-Host ''

Invoke-QuitCase -Name 'control: workbook only' -Setup {
    param($excel, $book)
    $null = $book.Worksheets.Item(1)
}

Invoke-QuitCase -Name 'table: ListObjects.Add' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $range = $sheet.Range('A1:C4')
    $null = $sheet.ListObjects.Add(1, $range, $null, 1)
}

Invoke-QuitCase -Name 'table + hidden column' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $sheet.Columns.Item(2).Hidden = $true
}

Invoke-QuitCase -Name 'table + quarter-point row height' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $sheet.Rows.Item(5).RowHeight = 0.25
}

Invoke-QuitCase -Name 'workbook defined name (Names.Add)' -Setup {
    param($excel, $book)
    $null = $book.Names.Add('GanttPlotAnchor', '=Sheet1!$A$1')
}

Invoke-QuitCase -Name 'table + Names.Add' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $null = $book.Names.Add('GanttPlotAnchor', '=Sheet1!$A$1')
}

Invoke-QuitCase -Name 'table + data validation' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    # xlValidateList = 3, xlValidAlertStop = 1, xlBetween = 1
    $sheet.Range('A2:A4').Validation.Add(3, 1, 1, 'a,b,c')
}

Invoke-QuitCase -Name 'table + shape' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $null = $sheet.Shapes.AddShape(1, 20.0, 40.0, 120.0, 30.0)
}

Invoke-QuitCase -Name 'second worksheet added' -Setup {
    param($excel, $book)
    $null = $book.Worksheets.Add()
}

Invoke-QuitCase -Name 'second worksheet + table + name' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $null = $book.Worksheets.Add()
    $null = $book.Names.Add('GanttPlotAnchor', '=Sheet1!$A$1')
}

Invoke-QuitCase -Name 'table + sheet protected then released' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $sheet.Protect()
    $sheet.Unprotect()
}

Write-Host ''
Write-Host '=== Is the table case deterministic? five repeats ==='
foreach ($attempt in 1..5) {
    Invoke-QuitCase -Name "table repeat $attempt" -Setup {
        param($excel, $book)
        $sheet = $book.Worksheets.Item(1)
        $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    }
}

Write-Host ''
Write-Host '=== Does forcing calculation before Quit change it? ==='
Invoke-QuitCase -Name 'table + Calculate before close' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $excel.Calculate()
}

Invoke-QuitCase -Name 'table + CalculateFull before close' -Setup {
    param($excel, $book)
    $sheet = $book.Worksheets.Item(1)
    $null = $sheet.ListObjects.Add(1, $sheet.Range('A1:C4'), $null, 1)
    $excel.CalculateFull()
}

Write-Host ''
Write-Host 'Done. Ordinary probe output; this script is disposable.'
