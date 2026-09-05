#requires -Version 7
<#
.SYNOPSIS
    Pester tests for pre-commit.ps1 and install-pre-commit.ps1
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $scriptPath = Join-Path $repoRoot 'scripts\pre-commit.ps1'
    $installScriptPath = Join-Path $repoRoot 'scripts\install-pre-commit.ps1'
}

Describe 'pre-commit.ps1' {
    It 'exists and is readable' {
        (Test-Path -LiteralPath $scriptPath) | Should -BeTrue
    }

    It 'runs only the fast deterministic gates, not full verify-quick' {
        $raw = Get-Content -LiteralPath $scriptPath -Raw
        # Strip the XML doc comment block (<# ... #>) so mentions in
        # .SYNOPSIS/.DESCRIPTION do not trip the negative match.
        $codeOnly = $raw -replace '(?s)<#.*?#>', ''
        $codeOnly | Should -Match 'check-cline-skills\.ps1'
        $codeOnly | Should -Match 'check-skill-summary\.ps1'
        $codeOnly | Should -Match 'check-status\.ps1'
        $codeOnly | Should -Match 'check-md-links\.ps1'
        # Must not INVOKE verify-quick.ps1 in the executable portion.
        $codeOnly | Should -Not -Match '(pwsh|\$\(|&)\s.*verify-quick\.ps1'
    }

    It 'fails fast with a non-zero exit on the first gate failure' {
        $content = Get-Content -LiteralPath $scriptPath -Raw
        $content | Should -Match 'exit \$LASTEXITCODE'
        $content | Should -Match 'Commit blocked'
    }
}

Describe 'install-pre-commit.ps1' {
    It 'exists and is readable' {
        (Test-Path -LiteralPath $installScriptPath) | Should -BeTrue
    }

    It 'invokes PowerShell 7 (pwsh) for the hook shim' {
        $content = Get-Content -LiteralPath $installScriptPath -Raw
        $content | Should -Match 'exec pwsh\.exe'
        $content | Should -Match 'core\.hooksPath'
    }

    It 'writes the shim as LF with UTF-8 without BOM' {
        $content = Get-Content -LiteralPath $installScriptPath -Raw
        $content | Should -Match 'New-Object System\.Text\.UTF8Encoding'
        $content | Should -Match '\$contentLf'
        $content | Should -Match 'WriteAllText'
    }
}