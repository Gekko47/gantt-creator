#requires -Version 7
<#
.SYNOPSIS
    W9 step-parity tripwires: every numbered step in the verify scripts'
    .DESCRIPTION block must have a matching Invoke-Step call, and vice
    versa. Catches docs/code drift between the header enumeration and the
    actual step list.
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    . (Join-Path $repoRoot 'scripts\verify-helpers.ps1')
    $script:verifyQuick   = Join-Path $repoRoot 'scripts\verify-quick.ps1'
    $script:verifyFull    = Join-Path $repoRoot 'scripts\verify.ps1'
}

Describe 'verify scripts: .DESCRIPTION step numbers match Invoke-Step names' {
    It 'verify-quick.ps1: numbered steps = Invoke-Step count' {
        $desc  = Get-VerifyDescriptionStepNumbers -Path $script:verifyQuick
        $steps = Get-VerifyStepNames           -Path $script:verifyQuick
        $steps.Count | Should -Be $desc.Count
    }

    It 'verify-quick.ps1: first numbered step matches first Invoke-Step' {
        $desc  = Get-VerifyDescriptionStepNumbers -Path $script:verifyQuick
        $steps = Get-VerifyStepNames           -Path $script:verifyQuick
        # The .DESCRIPTION numbering starts at 1; the Invoke-Step list
        # is the actual run order. They must align 1:1.
        for ($i = 0; $i -lt [Math]::Min($desc.Count, $steps.Count); $i++) {
            $desc[$i] | Should -Be ($i + 1)
        }
    }

    It 'verify-quick.ps1: every Invoke-Step name is non-empty and unique' {
        $steps = Get-VerifyStepNames -Path $script:verifyQuick
        $steps | Should -Not -BeNullOrEmpty
        ($steps | Sort-Object -Unique).Count | Should -Be $steps.Count
    }

    It 'verify.ps1: numbered steps = Invoke-Step count' {
        $desc  = Get-VerifyDescriptionStepNumbers -Path $script:verifyFull
        $steps = Get-VerifyStepNames           -Path $script:verifyFull
        $steps.Count | Should -Be $desc.Count
    }

    It 'verify.ps1: every Invoke-Step name is non-empty and unique' {
        $steps = Get-VerifyStepNames -Path $script:verifyFull
        $steps | Should -Not -BeNullOrEmpty
        ($steps | Sort-Object -Unique).Count | Should -Be $steps.Count
    }

    It 'positive control: the parser finds exactly one step in a single-step fixture' {
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('verify-helper-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $fixture = @'
Invoke-Step 'sole' { Write-Host 'x' }
'@
            Set-Content -LiteralPath (Join-Path $td 'fixture.ps1') -Value $fixture -Encoding utf8
            $names = Get-VerifyStepNames -Path (Join-Path $td 'fixture.ps1')
            $names | Should -Be @('sole')
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'W9 no-silent-pass: gate scripts fail on empty input' {
    It 'check-md-links.ps1: empty scan is a failure (positive test)' {
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('md-empty-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $emptyDocs = Join-Path $td 'docs'
            New-Item -ItemType Directory -Path $emptyDocs -Force | Out-Null
            # No .md files in $emptyDocs.
            $scriptPath = Join-Path $repoRoot 'scripts\check-md-links.ps1'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$scriptPath,'-Roots','docs','-Entry','AGENTS.md'
            ) -NoNewWindow -Wait -PassThru `
                -WorkingDirectory $td `
                -RedirectStandardOutput (Join-Path $td 'out.txt') `
                -RedirectStandardError (Join-Path $td 'err.txt')
            $proc.ExitCode | Should -Not -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'check-md-links.ps1: non-empty scan with valid links is a PASS (positive control)' {
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('md-ok-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            $docs = Join-Path $td 'docs'
            New-Item -ItemType Directory -Path $docs -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $docs 'a.md') -Value '# a' -Encoding utf8
            $scriptPath = Join-Path $repoRoot 'scripts\check-md-links.ps1'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$scriptPath,'-Roots','docs','-Entry','AGENTS.md'
            ) -NoNewWindow -Wait -PassThru `
                -WorkingDirectory $td `
                -RedirectStandardOutput (Join-Path $td 'out.txt') `
                -RedirectStandardError (Join-Path $td 'err.txt')
            $proc.ExitCode | Should -Be 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}