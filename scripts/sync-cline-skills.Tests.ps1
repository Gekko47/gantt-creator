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

This is a test canonical source for sync-cline-skills.

It has more than one line so we can also see the summary truncation.
'@ | Set-Content -LiteralPath $fixtureDoc -Encoding utf8

            $syncBody = Get-Content -LiteralPath $script:scriptPath -Raw
            $fixtureMap = @'
$map = @{
    '99-FIXTURE.md' = @{ Name = '99-fixture'; Description = 'Test skill.' }
}
'@
            $syncBody = $syncBody -replace '(?s)\$map = @\{.*?\}(?=\s*foreach)', ($fixtureMap + "`n")
            $script:harnessSync = Join-Path $script:tempRoot 'sync-cline-skills.ps1'
            $syncBody | Set-Content -LiteralPath $script:harnessSync -Encoding utf8
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exits 0 and regenerates SKILL.md + references.md on a valid fixture' {
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
            Test-Path -LiteralPath (Join-Path $script:skills '99-fixture/references.md') | Should -BeTrue

            # The first summary line of the canonical source is present
            # in the regenerated SKILL.md (observable output, not a
            # string match against the script source).
            $skillBody = Get-Content -LiteralPath (Join-Path $script:skills '99-fixture/SKILL.md') -Raw
            $skillBody | Should -Match 'Fixture document'
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
    }
}