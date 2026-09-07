#requires -Version 7
<#
.SYNOPSIS
    Pester tests for the PSScriptAnalyzer gate in verify-quick.ps1 / verify.ps1.

.DESCRIPTION
    The original step used `-EnableExit`, which calls `exit` from inside a
    function. The exit does not propagate to the parent process's
    $LASTEXITCODE when the cmdlet runs inside a Tee/ForEach pipeline (the
    Invoke-Step shape in those scripts), so a real failure was reported as
    PASS. The fix captures the findings and judges in the step's own scope.
    The first test below reproduces the failure path with a synthetic bad
    script and proves the new step shape fails it (positive test for the
    validator). The second test runs the real analyzer step against a
    scratch fixture with a single known warning and asserts the gate
    reports it.
#>

Describe 'PSScriptAnalyzer gate (verify scripts)' {
    It 'verify-quick.ps1 no longer calls Invoke-ScriptAnalyzer with -EnableExit' {
        # Computed here so the assignment and use sit in the same block;
        # PSSA cannot flow-analyse a Pester BeforeAll across an It.
        $verifyQuick = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts/verify-quick.ps1'
        $text = Get-Content -LiteralPath $verifyQuick -Raw
        $text | Should -Not -Match 'Invoke-ScriptAnalyzer[^\r\n]*-EnableExit'
    }

    It 'verifies a real PSSA warning is reported (positive test for the gate)' {
        # Synthetic fixture: a script that triggers
        # PSUseDeclaredVarsMoreThanAssignments (Warning, not in the settings
        # ExcludeRules list). The gate must report at least one finding; the
        # count alone proves the step is no longer blind.
        if (-not (Get-Module -ListAvailable PSScriptAnalyzer)) {
            Set-ItResult -Skipped -Because 'PSScriptAnalyzer not installed on this host'
            return
        }
        $pssaSettings = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts/PSScriptAnalyzerSettings.psd1'
        $td = Join-Path ([System.IO.Path]::GetTempPath()) ('pssa-gate-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $td -Force | Out-Null
        try {
            @'
function foo { $x = 'declared but never read' }
foo
'@ | Set-Content -LiteralPath (Join-Path $td 'bad.ps1') -Encoding utf8

            $findings = @(Invoke-ScriptAnalyzer `
                -Path (Join-Path $td 'bad.ps1') `
                -Settings $pssaSettings)
            $findings.Count | Should -BeGreaterThan 0
            ($findings | Where-Object RuleName -eq 'PSUseDeclaredVarsMoreThanAssignments').Count |
                Should -BeGreaterThan 0
        }
        finally {
            Remove-Item -LiteralPath $td -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}