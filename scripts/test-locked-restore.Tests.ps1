#requires -Version 7
<#
.SYNOPSIS
    Pester tests for test-locked-restore.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $scriptPath = Join-Path $repoRoot 'scripts\test-locked-restore.ps1'
}

Describe 'test-locked-restore.ps1' {
    It 'exists and is readable' {
        (Test-Path $scriptPath) | Should -BeTrue
    }

    It 'searches from repo root ($PSScriptRoot\..) not parent of repo' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '\$PSScriptRoot\\\\\\.\\. '
    }

    It 'runs two locked-mode restores with --no-cache' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match '--locked-mode.*--no-cache'
    }

    It 'deletes obj directories before each run' {
        (Get-Content -LiteralPath $scriptPath -Raw) | Should -Match 'Remove-ObjDirectory'
    }
}