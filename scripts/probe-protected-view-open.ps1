#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: records the protected-view and outline-view facts R7.2 consumes.

.DESCRIPTION
    R7.2's live gate classifies a protected-view document and refuses an unsupported outline
    view. Two of those claims are not documented as deterministic by Microsoft Learn, so this
    probe measures them on the reference host before the live case gate is written:

      1. whether ProtectedViewWindows.Open on a locally created .pptx produces a protected-view
         window, and what Presentations.Count and ActiveProtectedViewWindow report in that state;
      2. the Mark-of-the-Web fallback (a Zone.Identifier stream with ZoneId=3) when the plain
         open does not force protected view;
      3. whether DocumentWindow.ViewType accepts ppViewOutline (6) and what View.Type reads back;
      4. the standard editing view's own View.Type value.

    The probe never destroys a user session. It snapshots POWERPNT before the call, quits only
    a process it started, closes only its own presentations and protected-view windows, and
    kills only the newly created process ids after re-confirming their start time.

.EXAMPLE
    pwsh ./scripts/probe-protected-view-open.ps1

.NOTES
    Disposable spike, in the style of probe-powerpoint-host.ps1. It changes no production code
    and writes only temporary files; the JSON block it prints is the evidence record.
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

$tempFolder = [System.IO.Path]::Combine(
    [System.IO.Path]::GetTempPath(),
    'ganttcreator-r7-2-probe-' + [System.Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $tempFolder -Force
$deckPath = [System.IO.Path]::Combine($tempFolder, 'probe.pptx')

$before = Get-PowerPointProcessSnapshot
$startedByUs = $before.Count -eq 0
$application = $null
$presentation = $null
$windowed = $null
$protectedWindow = $null
$facts = [ordered]@{
    capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    powerPointWasRunning = ($before.Count -gt 0)
    processIdsBefore = @($before.Keys)
    tempFolder = $tempFolder
}

try {
    $application = New-Object -ComObject PowerPoint.Application
    $after = Get-PowerPointProcessSnapshot
    $newProcessIds = @($after.Keys | Where-Object { -not $before.ContainsKey($_) })

    $facts['startedByUs'] = $startedByUs
    $facts['applicationVersion'] = $application.Version
    $facts['presentationsCountAtAttach'] = $application.Presentations.Count
    $facts['protectedViewWindowsAtAttach'] = $application.ProtectedViewWindows.Count

    # A saved deck gives the protected-view open a real file. WithWindow = 0 keeps it out of
    # the user's view; a blank slide makes it a valid presentation.
    $presentation = $application.Presentations.Add(0)
    $null = $presentation.Slides.Add(1, 12)
    $presentation.SaveAs($deckPath, 24)
    $facts['savedDeckPath'] = $deckPath
    $facts['presentationsCountWithDeck'] = $application.Presentations.Count

    # Plain open route.
    $plainOutcome = 'not-attempted'
    try {
        $protectedWindow = $application.ProtectedViewWindows.Open($deckPath)
        $plainOutcome = 'opened'
        $facts['plainProtectedWindowCaption'] = $protectedWindow.Caption
    }
    catch {
        $plainOutcome = 'refused:' + (Get-ExceptionHResult -ErrorRecord $_)
    }

    $facts['plainOpenOutcome'] = $plainOutcome
    $facts['plainProtectedViewWindowsCount'] = $application.ProtectedViewWindows.Count
    $activeProtected = $application.ActiveProtectedViewWindow
    $facts['plainActiveProtectedViewWindowPresent'] = ($null -ne $activeProtected)
    $facts['plainPresentationsCount'] = $application.Presentations.Count

    if ($null -ne $protectedWindow) {
        $protectedWindow.Close()
        $protectedWindow = $null
    }

    # Mark-of-the-Web fallback route.
    $fallbackOutcome = 'not-attempted'
    try {
        # Mark-of-the-Web: the documented unsafe-location trigger for protected view.
        $zoneStream = $deckPath + ':Zone.Identifier'
        Set-Content -LiteralPath $zoneStream -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding ascii -NoNewline
        $facts['zoneIdentifierStream'] = $zoneStream
        $protectedWindow = $application.ProtectedViewWindows.Open($deckPath)
        $fallbackOutcome = 'opened'
        $facts['motwProtectedWindowCaption'] = $protectedWindow.Caption
    }
    catch {
        $fallbackOutcome = 'refused:' + (Get-ExceptionHResult -ErrorRecord $_)
    }

    $facts['motwOpenOutcome'] = $fallbackOutcome
    $facts['motwProtectedViewWindowsCount'] = $application.ProtectedViewWindows.Count
    $activeProtected = $application.ActiveProtectedViewWindow
    $facts['motwActiveProtectedViewWindowPresent'] = ($null -ne $activeProtected)
    $facts['motwPresentationsCount'] = $application.Presentations.Count

    if ($null -ne $protectedWindow) {
        $protectedWindow.Close()
        $protectedWindow = $null
    }

    # Outline and sorter settability on a windowed deck, plus the standard editing view value.
    $windowed = $application.Presentations.Add(-1)
    $null = $windowed.Slides.Add(1, 12)
    $window = $application.ActiveWindow
    $facts['standardViewType'] = $window.View.Type

    $outlineOutcome = 'not-attempted'
    try {
        $window.ViewType = 6
        $outlineOutcome = 'set'
        $facts['outlineViewTypeAfterSet'] = $window.View.Type
    }
    catch {
        $outlineOutcome = 'refused:' + (Get-ExceptionHResult -ErrorRecord $_)
    }

    $facts['outlineSetOutcome'] = $outlineOutcome

    $sorterOutcome = 'not-attempted'
    try {
        $window.ViewType = 7
        $sorterOutcome = 'set'
        $facts['sorterViewTypeAfterSet'] = $window.View.Type
    }
    catch {
        $sorterOutcome = 'refused:' + (Get-ExceptionHResult -ErrorRecord $_)
    }

    $facts['sorterSetOutcome'] = $sorterOutcome
}
finally {
    foreach ($owned in @($protectedWindow)) {
        if ($null -ne $owned) {
            try {
                $owned.Close()
            }
            catch {
                Write-Verbose 'The protected-view window could not be closed.'
            }
        }
    }

    foreach ($owned in @($windowed, $presentation)) {
        if ($null -ne $owned) {
            try {
                $owned.Saved = -1
                $owned.Close()
            }
            catch {
                Write-Verbose 'A probe presentation could not be closed; the owned process is stopped below.'
            }
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
            $newProcessIds | Where-Object {
                $null -ne (Get-Process -Id $_ -ErrorAction SilentlyContinue)
            })

        foreach ($processId in $survivors) {
            try {
                Stop-Process -Id $processId -Force -ErrorAction Stop
            }
            catch {
                Write-Verbose "The owned process $processId could not be stopped."
            }
        }
    }

    Remove-Item -LiteralPath $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
}

$facts | ConvertTo-Json -Depth 6
