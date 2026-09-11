#requires -Version 7
<#
.SYNOPSIS
    W9 step-parity tripwires: every numbered step in the verify scripts'
    .DESCRIPTION block must have a matching Invoke-Step call, and vice
    versa. Catches docs/code drift between the header enumeration and the
    actual step list.
#>

BeforeAll {
    # Top-level BeforeAll variables are not visible to It blocks under
    # Pester 6, so hold everything in $script: scope (same pattern as
    # review-runbook.Tests.ps1 and the W14 PSSA lesson).
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
    . (Join-Path $script:repoRoot 'scripts\verify-helpers.ps1')
    $script:verifyQuick   = Join-Path $script:repoRoot 'scripts\verify-quick.ps1'
    $script:verifyFull    = Join-Path $script:repoRoot 'scripts\verify.ps1'
}

Describe 'verify scripts: .DESCRIPTION step numbers match Invoke-Step names' {
    It 'verify-quick.ps1: numbered steps = Invoke-Step count' {
        $desc  = Read-VerifyDescriptionStepNumber -Path $script:verifyQuick
        $steps = Read-VerifyStepName           -Path $script:verifyQuick
        $steps.Count | Should -Be $desc.Count
    }

    It 'verify-quick.ps1: .DESCRIPTION numbering is sequential from 1' {
        $desc = Read-VerifyDescriptionStepNumber -Path $script:verifyQuick
        # The .DESCRIPTION numbering starts at 1 and increments by 1.
        for ($i = 0; $i -lt $desc.Count; $i++) {
            $desc[$i] | Should -Be ($i + 1)
        }
    }

    It 'verify-quick.ps1: every Invoke-Step name is non-empty and unique' {
        $steps = Read-VerifyStepName -Path $script:verifyQuick
        $steps | Should -Not -BeNullOrEmpty
        ($steps | Sort-Object -Unique).Count | Should -Be $steps.Count
    }

    It 'verify-quick.ps1: ordered description labels match ordered Invoke-Step names' {
        # Parse each description label (the text after the step number) and
        # compare the resulting ordered labels directly with $steps. Equal
        # totals with different step meanings must fail this one-to-one
        # ordered identity check.
        $descLabels = Read-VerifyDescriptionStepLabel -Path $script:verifyQuick
        $steps = Read-VerifyStepName -Path $script:verifyQuick
        $descLabels.Count | Should -Be $steps.Count
        for ($i = 0; $i -lt $steps.Count; $i++) {
            $descLabels[$i] | Should -Be $steps[$i] -Because "description label $i ('$($descLabels[$i])') must match Invoke-Step name '$($steps[$i])' in order"
        }
    }

    It 'verify.ps1: numbered steps = Invoke-Step count' {
        $desc  = Read-VerifyDescriptionStepNumber -Path $script:verifyFull
        $steps = Read-VerifyStepName           -Path $script:verifyFull
        $steps.Count | Should -Be $desc.Count
    }

    It 'verify.ps1: .DESCRIPTION numbering is sequential from 1' {
        $desc = Read-VerifyDescriptionStepNumber -Path $script:verifyFull
        # The .DESCRIPTION numbering starts at 1 and increments by 1.
        for ($i = 0; $i -lt $desc.Count; $i++) {
            $desc[$i] | Should -Be ($i + 1)
        }
    }

    It 'verify.ps1: every Invoke-Step name is non-empty and unique' {
        $steps = Read-VerifyStepName -Path $script:verifyFull
        $steps | Should -Not -BeNullOrEmpty
        ($steps | Sort-Object -Unique).Count | Should -Be $steps.Count
    }

    It 'verify.ps1: ordered description labels match ordered Invoke-Step names' {
        # Parse each description label (the text after the step number) and
        # compare the resulting ordered labels directly with $steps. Equal
        # totals with different step meanings must fail this one-to-one
        # ordered identity check.
        $descLabels = Read-VerifyDescriptionStepLabel -Path $script:verifyFull
        $steps = Read-VerifyStepName -Path $script:verifyFull
        $descLabels.Count | Should -Be $steps.Count
        for ($i = 0; $i -lt $steps.Count; $i++) {
            $descLabels[$i] | Should -Be $steps[$i] -Because "description label $i ('$($descLabels[$i])') must match Invoke-Step name '$($steps[$i])' in order"
        }
    }

    It 'positive control: the parser finds exactly one step in a single-step fixture' {
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('verify-helper-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $fixture = @'
Invoke-Step 'sole' { Write-Host 'x' }
'@
            Set-Content -LiteralPath (Join-Path $td 'fixture.ps1') -Value $fixture -Encoding utf8
            $names = Read-VerifyStepName -Path (Join-Path $td 'fixture.ps1')
            $names | Should -Be @('sole')
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'W9 no-silent-pass: gate scripts fail on empty input' {
    It 'check-md-links.ps1: empty scan is a failure (positive test)' {
        # check-md-links.ps1 anchors its roots to Split-Path -Parent $PSScriptRoot
        # (the real scripts/ dir), not the caller's CWD. We pass
        # -Roots 'nonexistent-only' and -Entry 'nonexistent-only' so the
        # foreach over roots does continue for every entry (Test-Path
        # returns false) and $scanned stays at 0 -- the documented
        # no-silent-pass branch.
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('md-empty-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $scriptPath = Join-Path $script:repoRoot 'scripts\check-md-links.ps1'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$scriptPath,'-Roots','nonexistent-only','-Entry','nonexistent-only'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput (Join-Path $td 'out.txt') `
                -RedirectStandardError (Join-Path $td 'err.txt')
            $proc.ExitCode | Should -Not -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'check-md-links.ps1: non-empty scan with valid links is a PASS (positive control)' {
        # check-md-links.ps1 resolves roots against Split-Path -Parent $PSScriptRoot,
        # so to scan the fixture we must copy the script into the harness and run
        # the copy -- its $repoRoot then points at the temp dir. The harness must
        # create all default configured roots (docs/, .github/, AGENTS.md) because
        # the script now fails when a configured root is missing. We pass -Roots
        # 'docs' only and let -Entry use its default 'AGENTS.md' (which the
        # harness creates). Passing -Entry as a file path ('a.md') would treat
        # it as a scan root and fail because the script expects directories.
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('md-ok-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $docs = Join-Path $td 'docs'
            New-Item -ItemType Directory -Path $docs -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $docs 'a.md') -Value '# a' -Encoding utf8
            $harnessScripts = Join-Path $td 'scripts'
            New-Item -ItemType Directory -Path $harnessScripts -Force | Out-Null
            $harnessScript = Join-Path $harnessScripts 'check-md-links.ps1'
            Copy-Item -LiteralPath (Join-Path $script:repoRoot 'scripts\check-md-links.ps1') -Destination $harnessScript
            # Create the other default configured roots so the script's default
            # parameter values work without triggering missing-root failures.
            $githubDir = Join-Path $td '.github'
            New-Item -ItemType Directory -Path $githubDir -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $td 'AGENTS.md') -Value '# agents' -Encoding utf8
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$harnessScript,'-Roots','docs'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput (Join-Path $td 'out.txt') `
                -RedirectStandardError (Join-Path $td 'err.txt')
            $proc.ExitCode | Should -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}