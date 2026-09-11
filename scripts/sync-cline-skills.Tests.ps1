#requires -Version 7
<#
.SYNOPSIS
    Pester tests for sync-cline-skills.ps1
#>

Describe 'sync-cline-skills.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\sync-cline-skills.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    Context 'isolated execution (replaces source-text assertions)' {
        BeforeEach {
            $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:docs      = Join-Path $script:tempRoot 'docs'
            $script:skills    = Join-Path $script:tempRoot '.cline/skills'
            New-Item -ItemType Directory -Path $script:docs   -Force | Out-Null
            New-Item -ItemType Directory -Path $script:skills -Force | Out-Null

            # Build a harness copy of sync-cline-skills.ps1 with a
            # single-entry $map pointing at one fixture canonical source.
            $fixtureDoc = Join-Path $script:docs '99-FIXTURE.md'
            @'
# Fixture document

<!-- SKILL-SUMMARY:START -->
This is a test canonical source for sync-cline-skills.

Do not get wrong:
- The SKILL-SUMMARY block is extracted verbatim.
<!-- SKILL-SUMMARY:END -->

<!-- SKILL-TOOLS:START -->
- `dotnet_test` — run tests.
- `dotnet_build` — build the solution.
<!-- SKILL-TOOLS:END -->

It has more than one line so we can also see the summary truncation.
'@ | Set-Content -LiteralPath $fixtureDoc -Encoding utf8

            $syncBody = Get-Content -LiteralPath $script:scriptPath -Raw
            $fixtureMap = @'
$map = @{
    '99-FIXTURE.md' = @{ Name = '99-fixture'; Description = 'Test skill.' }
}
'@
            # Match only the $map block: from "$map = @{" to the first "}"
            # on its own line. The old lookahead form (\}\s*foreach) broke
            # when the sync script gained variable declarations between the
            # map and its first foreach loop — the lazy match swallowed
            # them and the harness script failed at runtime.
            $originalBody = $syncBody
            $syncBody = $syncBody -replace '(?ms)\$map = @\{.*?^\}', ($fixtureMap + "`n")
            if ($syncBody -eq $originalBody) {
                throw "Harness setup failed: `$map substitution pattern did not match; the unmodified production map would be written to the harness script."
            }
            $script:harnessSync = Join-Path $script:tempRoot 'sync-cline-skills.ps1'
            $syncBody | Set-Content -LiteralPath $script:harnessSync -Encoding utf8
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exits 0 and regenerates SKILL.md on a valid fixture with SKILL-SUMMARY block' {
            $outFile = Join-Path $script:tempRoot 'out-ok.txt'
            $errFile = Join-Path $script:tempRoot 'err-ok.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $stdout = Get-Content -LiteralPath $outFile -Raw

            $proc.ExitCode | Should -Be 0
            $stdout | Should -Match 'regenerated 1 skills'
            Test-Path -LiteralPath (Join-Path $script:skills '99-fixture/SKILL.md')    | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $script:skills '99-fixture/references.md') | Should -BeFalse

            # The SKILL-SUMMARY block content is present in the regenerated SKILL.md
            $skillBody = Get-Content -LiteralPath (Join-Path $script:skills '99-fixture/SKILL.md') -Raw
            $skillBody | Should -Match 'The SKILL-SUMMARY block is extracted verbatim'
            # The footer points at the canonical source, not at a references.md duplicate
            $skillBody | Should -Match 'Full reference:.*docs/99-FIXTURE.md'
            $skillBody | Should -Not -Match 'references.md'
        }

        It 'exits non-zero when a canonical source listed in $map is missing' {
            # Delete the fixture canonical source so the validator that
            # walks $map.Keys triggers "Canonical source missing".
            Remove-Item -LiteralPath (Join-Path $script:docs '99-FIXTURE.md') -Force

            $outFile = Join-Path $script:tempRoot 'out-missing.txt'
            $errFile = Join-Path $script:tempRoot 'err-missing.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $stderr = Get-Content -LiteralPath $errFile -Raw
            $combined = (Get-Content -LiteralPath $outFile -Raw) + $stderr

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'Canonical source missing'
        }

        It 'exits non-zero when the canonical doc has no SKILL-SUMMARY block (fallback removed)' {
            # Regression guard for the removed truncation fallback: a mapped
            # doc without the required block must fail the sync with a clear
            # message, never silently degrade into a truncated summary.
            @'
# Fixture document without a summary block

Just a body, no SKILL-SUMMARY markers anywhere.
'@ | Set-Content -LiteralPath (Join-Path $script:docs '99-FIXTURE.md') -Encoding utf8

            $outFile = Join-Path $script:tempRoot 'out-nosummary.txt'
            $errFile = Join-Path $script:tempRoot 'err-nosummary.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'no SKILL-SUMMARY block'
        }
        It 'exits non-zero when the canonical doc has an empty SKILL-SUMMARY block' {
            # Positive failure-path test for the empty-summary validator:
            # whitespace-only summary must fail, never silently yield empty.
            @'
# Fixture document with an empty summary block

<!-- SKILL-SUMMARY:START -->

<!-- SKILL-SUMMARY:END -->

Body content follows.
'@ | Set-Content -LiteralPath (Join-Path $script:docs '99-FIXTURE.md') -Encoding utf8

            $outFile = Join-Path $script:tempRoot 'out-emptysummary.txt'
            $errFile = Join-Path $script:tempRoot 'err-emptysummary.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'empty SKILL-SUMMARY block'
        }
        It 'exits non-zero when the canonical doc has no SKILL-TOOLS block' {
            # Regression guard: a mapped doc without the required
            # SKILL-TOOLS block must fail the sync with a clear message.
            @'
# Fixture document without a tools block

<!-- SKILL-SUMMARY:START -->
This is a test canonical source for sync-cline-skills.
<!-- SKILL-SUMMARY:END -->

Just a body, no SKILL-TOOLS markers anywhere.
'@ | Set-Content -LiteralPath (Join-Path $script:docs '99-FIXTURE.md') -Encoding utf8

            $outFile = Join-Path $script:tempRoot 'out-notools.txt'
            $errFile = Join-Path $script:tempRoot 'err-notools.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'no SKILL-TOOLS block'
        }
        It 'exits non-zero when the canonical doc has an empty SKILL-TOOLS block' {
            # Positive failure-path test for the empty-tools validator:
            # whitespace-only tools must fail, never silently yield empty.
            @'
# Fixture document with an empty tools block

<!-- SKILL-SUMMARY:START -->
This is a test canonical source for sync-cline-skills.
<!-- SKILL-SUMMARY:END -->

<!-- SKILL-TOOLS:START -->

<!-- SKILL-TOOLS:END -->

Body content follows.
'@ | Set-Content -LiteralPath (Join-Path $script:docs '99-FIXTURE.md') -Encoding utf8

            $outFile = Join-Path $script:tempRoot 'out-emptytools.txt'
            $errFile = Join-Path $script:tempRoot 'err-emptytools.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessSync,'-DocsRoot',$script:docs,'-SkillsRoot',$script:skills
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'empty SKILL-TOOLS block'
        }
    }
}