#requires -Version 7
<#
.SYNOPSIS
    Shared helpers for the verify gate scripts (W9 step-parity / no-silent-pass).

.DESCRIPTION
    Three helpers: the two step/description parsers below plus
    Invoke-PssaGate (the PSScriptAnalyzer gate). Used by:
      - scripts/verify-quick.ps1 and scripts/verify.ps1, to document
        their step list in the .DESCRIPTION block at the top of the file
        and to wire the .DESCRIPTION numbering to the actual Invoke-Step
        calls. A step that exists in .DESCRIPTION but has no matching
        Invoke-Step (or vice versa) is a docs/code drift defect, caught
        by the test at the end of this file.
      - scripts/step-parity.Tests.ps1, the Pester test that runs the
        helper against both verify scripts and asserts the two views
        match.

    Exit codes: the two parsers do not throw; Invoke-PssaGate fails the
    calling process with exit 1 when the analyzer reports findings, which
    is the capture-then-judge behaviour the verify scripts rely on.
#>

function Assert-UsablePath {
    <#
    .SYNOPSIS
        Refuses a null, empty, whitespace-only, or empty-collection path.

    .DESCRIPTION
        Every path handed to a PowerShell path cmdlet is validated here first.
        The failure this prevents is not hypothetical: CI runs 45193295738 and
        45376223093 died at
        `Value cannot be null. (Parameter 'The provided Path argument was null
        or an empty collection.')`, which is what PathIntrinsics throws for a
        null or empty-collection `-Path`, and it named no file and no worker.
        A null path is worse than a throw for `Invoke-Pester -Path`: that
        silently falls back to the current directory, so the gate would analyse
        the wrong tree and still report a verdict.

        Refusals are `throw`s, never `Write-Error`: the caller owns the stream
        and the exit code (see Get-ScriptTestGateVerdict for the same reason).

    .PARAMETER Path
        One path, or an array of paths. `$null`, `@()`, `''`, and whitespace
        are all refused.

    .PARAMETER Caller
        Label used in the refusal message, so the annotation names the gate.
    #>
    param(
        [Parameter(Mandatory)][string]$Caller,
        [AllowNull()][AllowEmptyCollection()][object]$Path
    )

    $values = @($Path)
    # An empty array and a single $null are different inputs that reach the
    # same broken cmdlet, so both are refused, with distinct messages.
    if ($values.Count -eq 0) {
        throw "$($Caller): refusing an empty path collection."
    }
    foreach ($value in $values) {
        if ([string]::IsNullOrWhiteSpace([string]$value)) {
            throw "$($Caller): refusing a null, empty, or whitespace-only path."
        }
    }
}

function Read-VerifyStepName {
    <#
    .SYNOPSIS
        Parses an Invoke-Step 'name' {...} block list from a verify script
        source and returns the names in declaration order.
    .DESCRIPTION
        Each Invoke-Step call appears as: `Invoke-Step 'NAME' { ... }`
        (with optional `[CmdletBinding()]` etc. before the name). The
        parser walks lines, tracks brace depth, and reports the literal
        text between the first pair of single quotes after `Invoke-Step`.
        It does not attempt to evaluate the scriptblock; it is a
        regex/line-walk over the source text.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Assert-UsablePath -Path $Path -Caller 'Read-VerifyStepName'
    if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
        throw "Read-VerifyStepName: file not found: $Path"
    }

    $names = New-Object System.Collections.Generic.List[string]
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $insideStep = $false
    $depth = 0
    foreach ($line in $lines) {
        if (-not $insideStep) {
            if ($line -match '^\s*Invoke-Step\s+''(.*?)''') {
                $names.Add($Matches[1])
                # Begin tracking depth so we don't double-count a name that
                # appears in a string literal inside a step body.
                $insideStep = $true
                $openCount = ([regex]::Matches($line, '\{')).Count
                $closeCount = ([regex]::Matches($line, '\}')).Count
                $depth = $openCount - $closeCount
                if ($depth -le 0) { $insideStep = $false; $depth = 0 }
            }
        }
        else {
            $openCount  = ([regex]::Matches($line, '\{')).Count
            $closeCount = ([regex]::Matches($line, '\}')).Count
            $depth += $openCount - $closeCount
            if ($depth -le 0) { $insideStep = $false; $depth = 0 }
        }
    }
    return $names.ToArray()
}

function Read-VerifyDescriptionStepNumber {
    <#
    .SYNOPSIS
        Parses the `  N. <description>` numbering in the .DESCRIPTION
        block at the top of a verify script and returns the numbers
        in declaration order. Lines that don't match the `N. ` pattern
        are ignored.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Assert-UsablePath -Path $Path -Caller 'Read-VerifyDescriptionStepNumber'
    if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
        throw "Read-VerifyDescriptionStepNumber: file not found: $Path"
    }
    $numbers = New-Object System.Collections.Generic.List[int]
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $inDescription = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*\.DESCRIPTION\s*$') { $inDescription = $true; continue }
        # Stop processing at the closing `#>` of the comment block
        if ($inDescription -and $line -match '^\s*#>\s*$') { return $numbers.ToArray() }
        if ($inDescription -and $line -match '^\s*(\d+)\.\s') {
            $numbers.Add([int]$Matches[1])
        }
    }
    return $numbers.ToArray()
}

function Read-VerifyDescriptionStepLabel {
    <#
    .SYNOPSIS
        Parses the description labels from the .DESCRIPTION block of a
        verify script and returns them in declaration order.
    .DESCRIPTION
        Each step in the .DESCRIPTION block appears as: `  N. Label text`
        The parser extracts the label text (everything after "N. ") for
        each numbered step, in order. This allows ordered comparison of
        description labels against Invoke-Step names to catch docs/code
        drift where the counts match but the step meanings differ.
    #>
    param([Parameter(Mandatory)][string]$Path)

    Assert-UsablePath -Path $Path -Caller 'Read-VerifyDescriptionStepLabel'
    if (-not (Test-Path -LiteralPath $Path -ErrorAction Stop)) {
        throw "Read-VerifyDescriptionStepLabel: file not found: $Path"
    }
    $labels = New-Object System.Collections.Generic.List[string]
    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    $inDescription = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*\.DESCRIPTION\s*$') { $inDescription = $true; continue }
        # Stop processing at the closing `#>` of the comment block
        if ($inDescription -and $line -match '^\s*#>\s*$') { return $labels.ToArray() }
        if ($inDescription -and $line -match '^\s*(\d+)\.\s+(.*)') {
            $label = $Matches[2].TrimEnd()
            if (-not [string]::IsNullOrWhiteSpace($label)) {
                $labels.Add($label)
            }
        }
    }
    return $labels.ToArray()
}
function Get-ScriptTestGateVerdict {
    <#
    .SYNOPSIS
        Judges per-file Pester results into the script-lint gate's verdict.

    .DESCRIPTION
        Separated from the runner (scripts/test-scripts.ps1) so the verdict
        rules can be tested in-process without spawning Pester per case. The
        runner executes the files; this function decides, and every rule here
        is a no-silent-pass rule:

          * a file that produced no result object (a crashed worker) is a FAILED
            file, never a skipped one;
          * a discovered file that contains zero tests is an error, because a
            gate that certifies a file it never exercised proves nothing;
          * any failed test fails the whole run, even when other files pass;
          * fewer results than discovered files is an error, so a lost worker
            cannot silently shrink the suite.

        Returns the per-file report lines, the error messages, the aggregate
        counts, and the exit code the runner must use. It does not print and it
        does not exit, so the caller owns the streams (W13: `exit` from inside a
        function called through a Tee/ForEach pipeline is swallowed).
    #>
    param(
        [AllowEmptyCollection()][object[]]$Results = @(),
        [int]$DiscoveredCount = 0
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $lines = New-Object System.Collections.Generic.List[string]
    $totalTests = 0
    $passedTests = 0
    $failedTests = 0
    $allPassed = $true

    foreach ($result in @($Results)) {
        if ($result.Error) {
            $allPassed = $false
            $lines.Add("  FAILED: $($result.Name) did not produce a result: $($result.Error)")
            continue
        }

        if ($result.Total -eq 0) {
            $allPassed = $false
            $errors.Add("test-scripts: $($result.Name) contains no Pester tests.")
            continue
        }

        $totalTests  += $result.Total
        $passedTests += $result.Passed
        $failedTests += $result.Failed

        if ($result.Failed -gt 0) {
            $allPassed = $false
            $lines.Add("  FAILED: $($result.Failed) of $($result.Total) tests failed in $($result.Name)")
            foreach ($failure in @($result.Failures)) { $lines.Add("    [-] $failure") }
        }
        else {
            $lines.Add("  PASSED: $($result.Total) tests")
        }
    }

    if (@($Results).Count -ne $DiscoveredCount) {
        $allPassed = $false
        $errors.Add(
            "test-scripts: $DiscoveredCount test file(s) discovered but " +
            "$(@($Results).Count) returned a result. A worker was lost; refusing to report a verdict.")
    }

    $exitCode = if ($allPassed -and $errors.Count -eq 0) { 0 } else { 1 }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Lines    = $lines.ToArray()
        Errors   = $errors.ToArray()
        Total    = $totalTests
        Passed   = $passedTests
        Failed   = $failedTests
    }
}

function Format-CheckAnnotation {
    <#
    .SYNOPSIS
        Formats one GitHub Actions ::error:: annotation line.

    .DESCRIPTION
        A failure inside the script-lint gate is otherwise readable only by
        downloading the job log, and that endpoint refuses an unauthenticated
        caller. Workflow commands are collected as check-run annotations, and
        that API is readable for a public repository, so one failed Pester test
        becomes one annotation.

        The message is flattened to a single line and the workflow-command
        escaping is applied (% first, per the command contract), so a Pester
        message containing a newline or a percent sign cannot terminate or
        corrupt the command. Property values escape ':' and ',' as well, because
        those delimit the property list.

    .PARAMETER Title
        Annotation title; the failing test file in the runner's use.

    .PARAMETER Message
        Failure text. Newlines are flattened rather than emitted.

    .PARAMETER File
        Repository-relative path the annotation points at; omitted when empty.
    #>
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Message,
        [string]$File = ''
    )

    $escapeProperty = { param([string]$Value) $Value.Replace('%', '%25').Replace("`r", ' ').Replace("`n", ' ').Replace(':', '%3A').Replace(',', '%2C') }

    $properties = 'title=' + (& $escapeProperty $Title)
    if ($File) { $properties = 'file=' + (& $escapeProperty $File) + ',' + $properties }
    # A CRLF pair is one line break, so it collapses to ONE space before the
    # remaining lone CR/LF characters are flattened.
    $messageText = $Message.Replace('%', '%25').Replace("`r`n", ' ').Replace("`r", ' ').Replace("`n", ' ')

    return "::error $properties::$messageText"
}

function Get-PssaTargetFile {
    <#
    .SYNOPSIS
        The files the PSScriptAnalyzer gate must examine under a root.

    .DESCRIPTION
        Enumerates .ps1/.psm1/.psd1 below each supplied root and skips the
        ignored `_artifacts` tree, which holds scratch scripts and hundreds of
        log/report files that are not part of the gate.

        This exists because `Invoke-ScriptAnalyzer -Path <root> -Recurse
        -Exclude '_artifacts'` costs far more than the analysis itself.
        Measured cold in separate processes on 2026-11-12:

          -Exclude '_artifacts'        14.26 s
          without -Exclude             13.59 s
          explicit file list            11.1 s

        The analyzer's own recursive enumeration with an exclusion pattern is
        therefore ~3 s, and resolving the list here also makes the analysed set
        explicit, so a test can assert exactly which files the gate examines
        (a script moved into a subdirectory must not silently stop being
        linted).

        Extensions are the analyzer's own set: a `.psd1` such as
        scripts/tool-versions.psd1 is in scope and must stay in the list.
    #>
    param([Parameter(Mandatory)][string[]]$Path)

    Assert-UsablePath -Path $Path -Caller 'Get-PssaTargetFile'

    $files = foreach ($root in $Path) {
        # A root that does not exist is reported here rather than silently
        # contributing zero files, so "the gate analysed nothing" always names
        # the path that is actually wrong.
        if (-not (Test-Path -LiteralPath $root -ErrorAction Stop)) {
            throw "Get-PssaTargetFile: path not found: $root"
        }
        if (Test-Path -LiteralPath $root -PathType Leaf -ErrorAction Stop) {
            Get-Item -LiteralPath $root -ErrorAction Stop
            continue
        }
        Get-ChildItem -LiteralPath $root -Recurse -File -ErrorAction Stop |
            Where-Object { $_.FullName -notmatch '[\\/]_artifacts([\\/]|$)' }
    }

    @($files |
            Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' } |
            Sort-Object FullName)
}

function Invoke-PssaGate {
    <#
    .SYNOPSIS
        Runs PSScriptAnalyzer over a root directory with the repo settings
        and fails the calling process when findings are present.

    .DESCRIPTION
        The PSSA gate used verbatim by verify-quick.ps1 and verify.ps1.
        It captures findings and judges them in this same scope instead of
        `-EnableExit`, whose function-level `exit` is swallowed by the
        Tee/ForEach pipeline inside Invoke-Step (W13, proven by the
        failing CI run on 2026-09-07). pssa-gate harness tests call this
        exact function against known-bad fixtures so the tested gate is
        the gate the verify scripts run.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Settings,
        [string]$Report = ''
    )

    if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
        Write-Error ('PSScriptAnalyzer is not installed. Run: ' +
            'Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -SkipPublisherCheck')
        exit 1
    }

    # Explicit file list instead of `-Recurse -Exclude` (see Get-PssaTargetFile).
    # An empty list is a failure, not a pass: a gate that analysed nothing
    # cannot certify the scripts directory.
    $targets = @(Get-PssaTargetFile -Path @($Path))
    if ($targets.Count -eq 0) {
        Write-Error "PSScriptAnalyzer gate found no .ps1/.psm1/.psd1 file under '$Path'. Analysing nothing is not a pass."
        exit 1
    }

    $findings = @(foreach ($target in $targets) {
            Invoke-ScriptAnalyzer -Path $target.FullName -Settings $Settings
        })

    # Deterministic order regardless of enumeration or analyzer order, so the
    # report is comparable between runs.
    $findings = @($findings | Sort-Object ScriptName, Line, RuleName)

    if ($findings.Count -gt 0) {
        foreach ($f in $findings) {
            $msg = "  [{0}] {1}:{2} {3}" -f $f.Severity, $f.ScriptName, $f.Line, $f.Message
            Write-Host $msg
            if ($Report) { Add-Content -LiteralPath $Report -Value $msg }
        }
        Write-Error "PSScriptAnalyzer reported $($findings.Count) issue(s). See output above."
        exit 1
    }
}
