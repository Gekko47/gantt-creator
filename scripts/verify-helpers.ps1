#requires -Version 7
<#
.SYNOPSIS
    Shared helpers for the verify gate scripts (W9 step-parity / no-silent-pass).

.DESCRIPTION
    Two helpers, both used by:
      - scripts/verify-quick.ps1 and scripts/verify.ps1, to document
        their step list in the .DESCRIPTION block at the top of the file
        and to wire the .DESCRIPTION numbering to the actual Invoke-Step
        calls. A step that exists in .DESCRIPTION but has no matching
        Invoke-Step (or vice versa) is a docs/code drift defect, caught
        by the test at the end of this file.
      - scripts/step-parity.Tests.ps1, the Pester test that runs the
        helper against both verify scripts and asserts the two views
        match.

    Exit codes: none of the helpers throw; the verify scripts wrap them
    in their own Invoke-Step.
#>

function Get-VerifyStepNames {
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
        throw "Get-VerifyStepNames: file not found: $Path"
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

function Get-VerifyDescriptionStepNumbers {
    <#
    .SYNOPSIS
        Parses the `  N. <description>` numbering in the .DESCRIPTION
        block at the top of a verify script and returns the numbers
        in declaration order. Lines that don't match the `N. ` pattern
        are ignored.
    #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Get-VerifyDescriptionStepNumbers: file not found: $Path"
    }
    $numbers = New-Object System.Collections.Generic.List[int]
    $lines = Get-Content -LiteralPath $Path
    $inDescription = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*\.DESCRIPTION\s*$') { $inDescription = $true; continue }
        if ($inDescription -and $line -match '^\s*\.SYNOPSIS\s*$') { break }
        if ($inDescription -and $line -match '^\s*(\d+)\.\s') {
            $numbers.Add([int]$Matches[1])
        }
    }
    return $numbers.ToArray()
}
