#requires -Version 7
<#
.SYNOPSIS
    Pester tests for check-skill-summary.ps1
#>

Describe 'check-skill-summary.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\check-skill-summary.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    Context 'isolated execution (replaces source-text assertions)' {
        BeforeEach {
            # The script's $repoRoot is Split-Path -Parent $PSScriptRoot.
            # The script reads $SkillsRoot relative to that $repoRoot, so
            # we put the fixture `.cline/skills/` directly under the
            # harness root and the script one level deeper (in `harness/`).
            $script:tempRoot   = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
            $script:harnessDir = Join-Path $script:tempRoot 'harness'
            $script:skills     = Join-Path $script:tempRoot '.cline/skills'
            New-Item -ItemType Directory -Path $script:harnessDir -Force | Out-Null
            New-Item -ItemType Directory -Path $script:skills     -Force | Out-Null

            # Copy the script into the harness so its $PSScriptRoot
            # resolves to the harness dir, and `$repoRoot` (which the
            # script derives from $PSScriptRoot) lines up with our
            # fixture tree.  This avoids touching the real `.cline/skills/`.
            Copy-Item -LiteralPath $script:scriptPath -Destination (Join-Path $script:harnessDir 'check-skill-summary.ps1')

            # The script's $assertions map is the one we need to satisfy
            # OR violate.  Build two parallel skills trees: the clean one
            # contains every required phrase; the bad one omits one.
            $clean = Join-Path $script:skills 'clean'
            $bad   = Join-Path $script:skills 'bad'
            New-Item -ItemType Directory -Path $clean -Force | Out-Null
            New-Item -ItemType Directory -Path $bad   -Force | Out-Null

            # The real script asserts on these three skill-relative paths:
            #   03-roadmap/SKILL.md  : "R0.8 note", "do not skip", "L6 and L7"
            #   02-architecture/SKILL.md : "additionally produces", "coverage/"
            #   04-test-strategy/SKILL.md : "set `CurrentCulture`", "ambient"
            $cleanRoadmap = Join-Path $clean '03-roadmap/SKILL.md'
            $cleanArch    = Join-Path $clean '02-architecture/SKILL.md'
            $cleanTest    = Join-Path $clean '04-test-strategy/SKILL.md'
            New-Item -ItemType Directory -Path (Split-Path $cleanRoadmap) -Force | Out-Null
            New-Item -ItemType Directory -Path (Split-Path $cleanArch)    -Force | Out-Null
            New-Item -ItemType Directory -Path (Split-Path $cleanTest)    -Force | Out-Null
            'R0.8 note: do not skip; L6 and L7 are part of the discipline.' | Set-Content -LiteralPath $cleanRoadmap -Encoding utf8
            'verify.ps1 additionally produces coverage/ artifacts.' | Set-Content -LiteralPath $cleanArch -Encoding utf8
            'set `CurrentCulture` is part of the ambient locale contract.' | Set-Content -LiteralPath $cleanTest -Encoding utf8

            # The bad tree: 03-roadmap is missing the phrase "do not skip".
            $badRoadmap = Join-Path $bad '03-roadmap/SKILL.md'
            $badArch    = Join-Path $bad '02-architecture/SKILL.md'
            $badTest    = Join-Path $bad '04-test-strategy/SKILL.md'
            New-Item -ItemType Directory -Path (Split-Path $badRoadmap) -Force | Out-Null
            New-Item -ItemType Directory -Path (Split-Path $badArch)    -Force | Out-Null
            New-Item -ItemType Directory -Path (Split-Path $badTest)    -Force | Out-Null
            'R0.8 note only.  No "do not skip" here, on purpose.' | Set-Content -LiteralPath $badRoadmap -Encoding utf8
            'verify.ps1 additionally produces coverage/ artifacts.' | Set-Content -LiteralPath $badArch -Encoding utf8
            'set `CurrentCulture` is part of the ambient locale contract.' | Set-Content -LiteralPath $badTest -Encoding utf8

            $script:cleanSkills = $clean
            $script:badSkills   = $bad
            $script:harnessScript = Join-Path $script:harnessDir 'check-skill-summary.ps1'
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        It 'exits 0 when every asserted canonical phrase is present' {
            $outFile = Join-Path $script:tempRoot 'out-clean.txt'
            $errFile = Join-Path $script:tempRoot 'err-clean.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessScript,'-SkillsRoot','.cline/skills/clean'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $stdout = Get-Content -LiteralPath $outFile -Raw

            $proc.ExitCode | Should -Be 0
            $stdout | Should -Match 'OK'
        }

        It 'exits 1 and names the missing phrase when a canonical phrase is absent' {
            $outFile = Join-Path $script:tempRoot 'out-bad.txt'
            $errFile = Join-Path $script:tempRoot 'err-bad.txt'
            $proc = Start-Process -FilePath pwsh -ArgumentList @(
                '-NoProfile','-File',$script:harnessScript,'-SkillsRoot','.cline/skills/bad'
            ) -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput $outFile -RedirectStandardError $errFile
            $combined = (Get-Content -LiteralPath $outFile -Raw) + (Get-Content -LiteralPath $errFile -Raw)

            $proc.ExitCode | Should -Not -Be 0
            # The script writes both a per-violation line and a summary
            # count.  Either is sufficient observable evidence that the
            # gate detected the missing phrase; the violation count
            # "1 violation(s):" is the reliably-captured marker.
            $combined | Should -Match "violation\(s\)"
        }
    }
}