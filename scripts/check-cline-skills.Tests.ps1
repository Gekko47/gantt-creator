#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-cline-skills.ps1
#>

Describe 'check-cline-skills.ps1' {
    BeforeAll {
        $script:repoRoot = Split-Path -Parent $PSScriptRoot
        $script:scriptPath = Join-Path $script:repoRoot 'scripts\check-cline-skills.ps1'
    }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    Context 'isolated execution (replaces source-text assertions)' {
        BeforeEach {
            $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:harness  = Join-Path $script:tempRoot 'scripts'
            $script:docs      = Join-Path $script:tempRoot 'docs'
            $script:skills    = Join-Path $script:tempRoot '.cline/skills'
            New-Item -ItemType Directory -Path $script:harness -Force | Out-Null
            New-Item -ItemType Directory -Path $script:docs    -Force | Out-Null
            New-Item -ItemType Directory -Path $script:skills  -Force | Out-Null

            # The gate does `git diff --name-only` and `git status --porcelain`
            # against docs/ and .cline/skills/.  In a non-repo both fail, so
            # the harness must be a real git repo with at least one commit
            # before any fixture is added (or the prime below would itself
            # be "dirty").
            git -C $script:tempRoot init -q | Out-Null
            git -C $script:tempRoot config user.email 'test@example.com' | Out-Null
            git -C $script:tempRoot config user.name  'test'              | Out-Null
            'init' | Set-Content -LiteralPath (Join-Path $script:tempRoot 'README.md') -Encoding utf8
            git -C $script:tempRoot add README.md | Out-Null
            git -C $script:tempRoot commit -q -m 'init' | Out-Null

            # One canonical fixture source so the sync runs in isolation
            # without needing all 7 real canonical files. It carries a
            # SKILL-SUMMARY block because the sync now fails on a doc
            # without one (the truncation fallback is removed).
            $fixtureDoc = Join-Path $script:docs '99-FIXTURE.md'
            @'
# Fixture document

<!-- SKILL-SUMMARY:START -->
This is a test canonical source.
<!-- SKILL-SUMMARY:END -->

<!-- SKILL-TOOLS:START -->
- `dotnet_test` — run tests.
- `dotnet_build` — build the solution.
<!-- SKILL-TOOLS:END -->

This is a test canonical source.
'@ | Set-Content -LiteralPath $fixtureDoc -Encoding utf8

            # Inject a single-entry $map into a copy of sync-cline-skills.ps1
            # so it can run inside the harness.
            $syncBody = Get-Content -LiteralPath (Join-Path $script:repoRoot 'scripts\sync-cline-skills.ps1') -Raw
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
            $syncBody = $syncBody -replace '(?ms)\$map = @\{.*?^\}', ($fixtureMap + "`n")
            $script:harnessSync = Join-Path $script:harness 'sync-cline-skills.ps1'
            $syncBody | Set-Content -LiteralPath $script:harnessSync -Encoding utf8

            # Copy check-cline-skills.ps1 unchanged into the harness; it
            # re-invokes the harness sync via -SkillsRoot.
            Copy-Item (Join-Path $script:repoRoot 'scripts\check-cline-skills.ps1') $script:harness

            # Prime the committed .cline/skills/ view so Phase 2 has a
            # committed view to byte-compare against.  Pass -DocsRoot
            # and -SkillsRoot as absolute paths so the sync does not
            # resolve them against the parent test CWD.
            pwsh -NoProfile -File $script:harnessSync -DocsRoot $script:docs -SkillsRoot $script:skills 2>$null | Out-Null

            # Commit the primed tree so the baseline is clean; per-test
            # `It` blocks may then deliberately introduce dirt or drift.
            # `-c core.autocrlf=false` keeps the committed bytes identical
            # to what the sync wrote (LF) so Phase 2's byte comparison
            # against a fresh regen is meaningful.
            git -C $script:tempRoot -c core.autocrlf=false add -A | Out-Null
            git -C $script:tempRoot -c core.autocrlf=false commit -q -m 'prime' | Out-Null
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exits 1 when docs/ has unstaged changes (Phase 1 dirty-tree detection)' {
            # Make the canonical fixture source dirty.  Phase 1 must fail
            # before Phase 2 even runs.
            $fixtureDoc = Join-Path $script:docs '99-FIXTURE.md'
            'local edit' | Add-Content -LiteralPath $fixtureDoc -Encoding utf8

            $outFile = Join-Path $script:tempRoot 'out-dirty.txt'
            $errFile = Join-Path $script:tempRoot 'err-dirty.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',(Join-Path $script:harness 'check-cline-skills.ps1')
            ) -NoNewWindow -Wait -PassThru `
                -WorkingDirectory $script:tempRoot `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'WORKING TREE DIRTY'
        }

        It 'exits 0 when both views are clean and in sync' {
            # Tree clean + temp-regen matches committed view.  Both
            # phases pass; observable contract is exit 0 + "in sync".
            $outFile = Join-Path $script:tempRoot 'out-clean.txt'
            $errFile = Join-Path $script:tempRoot 'err-clean.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',(Join-Path $script:harness 'check-cline-skills.ps1')
            ) -NoNewWindow -Wait -PassThru `
                -WorkingDirectory $script:tempRoot `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Be 0
            $combined | Should -Match 'in sync'
        }

        It 'exits 1 with DRIFT DETECTED when the committed skill view is stale (Phase 2)' {
            # Tree clean, but the committed SKILL.md was hand-edited to
            # differ from what a fresh regeneration would produce.
            $skillFile = Join-Path $script:skills '99-fixture/SKILL.md'
            'stale content that will not match the fresh sync output' | Set-Content -LiteralPath $skillFile -Encoding utf8

            # Commit the stale content so the working tree is clean
            # (Phase 1 must pass); only Phase 2 should fire on the
            # committed-vs-regenerated diff.
            git -C $script:tempRoot -c core.autocrlf=false add -A | Out-Null
            git -C $script:tempRoot -c core.autocrlf=false commit -q -m 'stale' | Out-Null

            $outFile = Join-Path $script:tempRoot 'out-drift.txt'
            $errFile = Join-Path $script:tempRoot 'err-drift.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',(Join-Path $script:harness 'check-cline-skills.ps1')
            ) -NoNewWindow -Wait -PassThru `
                -WorkingDirectory $script:tempRoot `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            $combined | Should -Match 'DRIFT DETECTED'
        }
    }
}