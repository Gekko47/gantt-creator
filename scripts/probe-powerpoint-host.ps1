#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: records the desktop PowerPoint facts Phase 7 consumes.

.DESCRIPTION
    R7.1 Step 0 folds the host probe into one evidence commit. The probe records what the
    typed adapter and the later rows must be built on, without adding production code:

      1. whether PowerPoint was running before the call, and the pre/post POWERPNT set;
      2. Presentations.Count, the active window, the view type and any slide-show window;
      3. a windowless presentation's slide count before and after adding a blank slide;
      4. the Presentation.ReadOnly fact for the test-owned presentation;
      5. the DisplayAlerts value, so R7.5 knows whether it must be saved and restored;
      6. the HRESULT produced by a missing member (an unsupported-build signal, never a
         transient retry).

    The production-scale ppPasteShape range and the three group-resize routes need the
    staged Excel group, so they are measured by the live R7.2/R7.3 gates and are not
    claimed here.

    The probe never destroys a user session. It snapshots POWERPNT before the call, quits
    only a process it started, closes only its own presentation, and kills only the newly
    created process ids after re-confirming their start time.

.EXAMPLE
    pwsh ./scripts/probe-powerpoint-host.ps1

.NOTES
    Disposable spike, in the style of probe-shaperange-bulk-style.ps1. It changes no
    production code and writes no file; the JSON block it prints is the evidence record.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PowerPointProcessSnapshot {
    $snapshot = [System.Collections.Generic.Dictionary[int, long]]::new()
    foreach ($process in Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) {
        try {
            $snapshot[$process.Id] = $process.StartTime.ToUniversalTime().Ticks
        }
        catch {
            Write-Verbose 'A PowerPoint process exited between enumeration and the read.'
        }
        finally {
            $process.Dispose()
        }
    }

    return $snapshot
}

function Get-ExceptionHResult {
    param([Parameter(Mandatory)] [System.Management.Automation.ErrorRecord] $ErrorRecord)

    $exception = $ErrorRecord.Exception
    while ($null -ne $exception.InnerException) {
        $exception = $exception.InnerException
    }

    return ('0x{0:X8}' -f $exception.HResult)
}

$before = Get-PowerPointProcessSnapshot
$after = $null
$startedByUs = $before.Count -eq 0
$application = $null
$presentation = $null
$createdPresentation = $false
$facts = [ordered]@{
    capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    powerPointWasRunning = ($before.Count -gt 0)
    processIdsBefore = @($before.Keys)
}

try {
    $application = New-Object -ComObject PowerPoint.Application
    $after = Get-PowerPointProcessSnapshot
    $newProcessIds = @($after.Keys | Where-Object { -not $before.ContainsKey($_) })

    $facts['startedByUs'] = $startedByUs
    $facts['processIdsAfter'] = @($after.Keys)
    $facts['newProcessIds'] = $newProcessIds
    $facts['applicationVersion'] = $application.Version
    $facts['presentationsCount'] = $application.Presentations.Count
    $facts['slideShowWindowsCount'] = $application.SlideShowWindows.Count
    $facts['displayAlerts'] = $application.DisplayAlerts

    $activeWindow = $application.ActiveWindow
    $facts['activeWindowPresent'] = ($null -ne $activeWindow)
    if ($null -ne $activeWindow) {
        $facts['activeViewType'] = $activeWindow.View.Type
    }

    # WithWindow = 0 (msoFalse): the probe's presentation stays out of the user's view.
    $presentation = $application.Presentations.Add(0)
    $createdPresentation = $true
    $facts['readOnly'] = $presentation.ReadOnly
    $facts['slidesCountBeforeAdd'] = $presentation.Slides.Count
    $slide = $presentation.Slides.Add(1, 12)
    $facts['slidesCountAfterAdd'] = $presentation.Slides.Count
    $facts['addedSlideIndex'] = $slide.SlideIndex

    try {
        $null = $application.GetType().InvokeMember(
            'GanttCreatorMissingMember',
            [System.Reflection.BindingFlags]::GetProperty,
            $null,
            $application,
            $null)
        $facts['missingMemberHResult'] = 'no-error'
    }
    catch {
        $facts['missingMemberHResult'] = Get-ExceptionHResult -ErrorRecord $_
    }
}
finally {
    if ($null -ne $presentation -and $createdPresentation) {
        try {
            $presentation.Saved = -1
            $presentation.Close()
        }
        catch {
            Write-Verbose 'The probe presentation could not be closed; the owned process is stopped below.'
        }
    }

    if ($null -ne $application) {
        if ($startedByUs) {
            try {
                $application.Quit()
            }
            catch {
                Write-Verbose 'Quit was refused; the owned-process sweep below is the fallback.'
            }
        }

        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($application)
    }

    if ($startedByUs) {
        $survivors = @(
            $facts['newProcessIds'] | Where-Object {
                $null -ne (Get-Process -Id $_ -ErrorAction SilentlyContinue)
            })

        if ($survivors.Count -gt 0) {
            Wait-Process -Id $survivors -Timeout 10 -ErrorAction SilentlyContinue
            foreach ($processId in $survivors) {
                $live = Get-Process -Id $processId -ErrorAction SilentlyContinue
                if ($null -eq $live) {
                    continue
                }

                # Only a process whose start time still matches the recorded one is ours.
                if ($null -ne $after
                    -and $after.ContainsKey($processId)
                    -and $live.StartTime.ToUniversalTime().Ticks -eq $after[$processId]) {
                    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
                }

                $live.Dispose()
            }
        }
    }
}

$facts | ConvertTo-Json -Depth 4
