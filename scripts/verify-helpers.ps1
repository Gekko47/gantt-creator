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

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Read-VerifyStepName: file not found: $Path"
    }

    $names = New-Object System.Collections.Generic.List[string]
    $lines = Get-Content -LiteralPath $Path
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

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Read-VerifyDescriptionStepNumber: file not found: $Path"
    }
    $numbers = New-Object System.Collections.Generic.List[int]
    $lines = Get-Content -LiteralPath $Path
    $inDescription = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*\.DESCRIPTION\s*$') { $inDescription = $true; continue }
        # Stop at the closing `#>` of the comment block, not at a later
        # .SYNOPSIS: comment trailers or code after the terminator must not
        # be scanned as description steps.
        if ($inDescription -and $line -match '^\s*#>\s*$') { break }
        if ($inDescription -and $line -match '^\s*(\d+)\.\s') {
            $numbers.Add([int]$Matches[1])
        }
    }
    return $numbers.ToArray()
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

    $findings = @(Invoke-ScriptAnalyzer -Path $Path -Recurse -Exclude '_artifacts' -Settings $Settings)
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
