#requires -Version 7
<#
.SYNOPSIS
    Pester tests for test-locked-restore.ps1

.DESCRIPTION
    Drives the script's helpers (Remove-ObjDirectory, Invoke-LockRestore)
    in a child pwsh with a stubbed `dotnet` that records its argv to a
    log file. The three assertions below observe the recorded trace and
    the resulting filesystem state directly -- they no longer rely on
    text-level regex matches against the script source.
#>

Describe 'test-locked-restore.ps1' {
    BeforeAll { $script:scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts\test-locked-restore.ps1' }

    It 'exists and is readable' {
        (Test-Path -LiteralPath $script:scriptPath) | Should -BeTrue
    }

    Context 'controlled execution' {
        BeforeAll {
            # Build a self-contained harness rooted at $script:tempRoot
            # that mirrors the real layout (the copy lives in a
            # `scripts\` subdir so its $PSScriptRoot\.. is $script:tempRoot
            # -- the real script's `scripts\` subdir has the same shape).
            # `dotnet` is shadowed by a shim on PATH that records each
            # invocation to $script:invocationLog and exits 0.
            function Write-LockedRestoreHarness {
                $script:tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())
                $script:scriptsDir = Join-Path $script:tempRoot 'scripts'
                $script:stubDir    = Join-Path $script:tempRoot 'dotnet-stub'
                $script:invocationLog = Join-Path $script:tempRoot 'invocation.log'
                $script:outsideRoot = Join-Path ([System.IO.Path]::GetTempPath()) ([System.Guid]::NewGuid())

                New-Item -ItemType Directory -Path $script:scriptsDir  -Force | Out-Null
                New-Item -ItemType Directory -Path $script:stubDir     -Force | Out-Null
                New-Item -ItemType Directory -Path $script:outsideRoot -Force | Out-Null
                '' | Set-Content -LiteralPath $script:invocationLog

                # Copy the script's helper definitions only. Boundaries
                # (top-level function defs only, after CRLF normalisation)
                # are structural rather than positional, so this stays
                # correct as the script grows. Defensive checks below turn
                # each failure mode into a clear message instead of an
                # opaque harness-wrapper error.
                $sourceText = Get-Content -LiteralPath $script:scriptPath -Raw

                $requiresMatch = [regex]::Match($sourceText, '#requires\s+-Version\s+(\d+)')
                if (-not $requiresMatch.Success) {
                    throw "$script:scriptPath must declare '#requires -Version N'; this test relies on that contract."
                }

                $parserType = [System.Management.Automation.Language.Parser]
                $tokens = $null
                $parseErrors = $null
                $ast = $parserType::ParseInput($sourceText, [ref]$tokens, [ref]$parseErrors)
                if ($parseErrors.Count -gt 0) {
                    throw "$script:scriptPath failed to parse: $($parseErrors -join '; ')"
                }

                $helperFuncs = $ast.FindAll(
                    { param($n)
                      $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                      $n.Parent -is [System.Management.Automation.Language.NamedBlockAst] -and
                      $n.Parent.Parent -is [System.Management.Automation.Language.ScriptBlockAst]
                    },
                    $true)

                if ($helperFuncs.Count -lt 2) {
                    throw "Expected to extract Remove-ObjDirectory and Invoke-LockRestore from $script:scriptPath but found $($helperFuncs.Count) top-level functions: $($helperFuncs.Name -join ', ')"
                }

                foreach ($name in 'Remove-ObjDirectory', 'Invoke-LockRestore') {
                    if ($helperFuncs.Name -notcontains $name) {
                        throw "Extracted helpers are missing '$name'. Found: $($helperFuncs.Name -join ', '). The script under test may have been refactored in a way this test cannot follow."
                    }
                }

                # The script's helpers reference $Solution (a script-scope
                # parameter). The copy must keep the param block too, or
                # the helpers see $Solution as $null and dotnet receives
                # the wrong argv.
                $paramBlockText = $null
                if ($ast.ParamBlock) {
                    $paramBlockText = $ast.ParamBlock.Extent.Text
                }
                if (-not $paramBlockText -or $paramBlockText -notmatch [regex]::Escape('[string]$Solution')) {
                    throw "$script:scriptPath must declare a [string] `$Solution parameter at script scope; the harness relies on it."
                }

                $helpersCode = "#requires -Version $($requiresMatch.Groups[1].Value)`n`n" +
                    ($paramBlockText -replace "`r`n", "`n") + "`n`n" +
                    "`$ErrorActionPreference = 'Stop'`n`n" +
                    (($helperFuncs | ForEach-Object {
                         ($_.Extent.Text -replace "`r`n", "`n") -replace "`r", "`n"
                     }) -join "`n`n")

                # Extract the script's own top-level statements from the same
                # parsed AST (the unnamed block of the ScriptBlockAst),
                # excluding the function definitions helpersCode already
                # carries. The harness replays exactly this sequence instead
                # of a hardcoded copy of it, so the test follows the script
                # if the cleanup/restore sequence ever changes.
                $mainBlock = $ast.FindAll(
                    { param($n)
                      $n -is [System.Management.Automation.Language.NamedBlockAst] -and
                      $n.Parent -is [System.Management.Automation.Language.ScriptBlockAst]
                    },
                    $true) | Select-Object -First 1
                if (-not $mainBlock) {
                    throw "$script:scriptPath has no top-level statement block; the harness cannot derive the execution sequence."
                }
                $topLevelStatements = @(
                    $mainBlock.Statements |
                        Where-Object { $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] }
                )
                if ($topLevelStatements.Count -lt 4) {
                    throw ("Expected the top-level cleanup/restore sequence in {0} but found {1} non-function statement(s)." -f $script:scriptPath, $topLevelStatements.Count)
                }
                $topLevelCode = ($topLevelStatements | ForEach-Object {
                    ($_.Extent.Text -replace "`r`n", "`n") -replace "`r", "`n"
                }) -join "`n`n"

                $script:copyScriptPath = Join-Path $script:scriptsDir 'test-locked-restore.ps1'
                Set-Content -LiteralPath $script:copyScriptPath -Value $helpersCode -Encoding utf8NoBOM

                # `dotnet` shim. The .cmd wrapper hands its argv to a
                # .ps1 stub so the captured argv is exact and free of
                # cmd-quoting noise. Each call appends a line of the
                # form `dotnet <args> nugget_packages=<set|unset>` to the
                # invocation log. The stub mirrors the real blocked-source
                # behaviour the script relies on: when NUGET_PACKAGES is
                # isolated to the empty control directory, package
                # acquisition fails (exit 1); otherwise restore
                # succeeds (exit 0).
                $script:dotnetStubPs1 = Join-Path $script:stubDir 'dotnet-impl.ps1'
                $stubPs1 = @"
`$ErrorActionPreference = 'Stop'
Add-Content -LiteralPath '$($script:invocationLog)' -Value ('dotnet ' + (`$args -join ' ') + ' nugget_packages=' + `$(if (`$env:NUGET_PACKAGES) { 'set' } else { 'unset' }))
if (`$env:NUGET_PACKAGES) { exit 1 }
exit 0
"@
                Set-Content -LiteralPath $script:dotnetStubPs1 -Value $stubPs1 -Encoding utf8NoBOM
                $stubCmd = Join-Path $script:stubDir 'dotnet.cmd'
                $cmdBody = "@echo off`r`npwsh -NoProfile -File `"$($script:dotnetStubPs1)`" %*"
                Set-Content -LiteralPath $stubCmd -Value $cmdBody -Encoding utf8NoBOM

                # obj/ inside the search root ($script:scriptsDir's
                # parent, $script:tempRoot -- exactly the script's
                # $PSScriptRoot\..). Must be deleted by the harness.
                $script:objInside = Join-Path $script:tempRoot 'obj'
                New-Item -ItemType Directory -Path $script:objInside -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $script:objInside 'inside.txt') -Value 'inside'

                # obj/ outside the search root -- must survive.
                $script:objOutside = Join-Path $script:outsideRoot 'obj'
                New-Item -ItemType Directory -Path $script:objOutside -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $script:objOutside 'outside.txt') -Value 'outside'

                # Harness wrapper. Dot-sources the copy (so the
                # script's helpers enter the wrapper's scope), captures
                # the original helper bodies, then installs recording
                # wrappers under the script's original helper names.
                # The recording wrappers log one line per call and
                # delegate to the captured body, so the helpers behave
                # exactly as the script defines them (including
                # $PSScriptRoot, which stays bound to the copy's
                # directory). The wrapper then drives the same sequence
                # as the script's top-level code.
                $script:harnessPath = Join-Path $script:tempRoot 'harness.ps1'
                $harnessBody = @"
#requires -Version 7
`$ErrorActionPreference = 'Stop'
`$env:PATH = '$($script:stubDir);' + `$env:PATH
# Clear any inherited NUGET_PACKAGES so the replayed sequence starts from
# a deterministic environment: only the script's own isolation control
# (run 3) sets it, keeping the single nugget_packages=set assertion stable
# when the parent shell exports the variable (e.g. a global NuGet cache
# folder).
Remove-Item Env:NUGET_PACKAGES -ErrorAction SilentlyContinue
. "$($script:copyScriptPath)"

# Capture the original helper bodies BEFORE we shadow the names.
`$originalRemoveObjBody = (Get-Command Remove-ObjDirectory).ScriptBlock
`$originalRestoreBody   = (Get-Command Invoke-LockRestore).ScriptBlock

# Recording wrappers. Each delegates to the captured body so the
# helper's own behavior (incl. `$PSScriptRoot` resolution, which
# stays bound to the copy's directory) is unchanged.
function Remove-ObjDirectory {
    [CmdletBinding(SupportsShouldProcess)]
    param()
    Add-Content -LiteralPath '$($script:invocationLog)' -Value 'Remove-ObjDirectory'
    & `$originalRemoveObjBody @args
}
function Invoke-LockRestore {
    param([string]`$Label, [string]`$ConfigFile, [switch]`$ExpectFailure)
    Add-Content -LiteralPath '$($script:invocationLog)' -Value ('Invoke-LockRestore ' + `$Label)
    & `$originalRestoreBody @PSBoundParameters
}

# The script's own top-level statements, extracted from its parsed AST
# above, so the harness always replays exactly the sequence the script
# defines (two clean restores, a denied-network control, an offline
# restore) instead of a hardcoded copy of it.
$topLevelCode
"@
                Set-Content -LiteralPath $script:harnessPath -Value $harnessBody -Encoding utf8NoBOM
            }

            # Run the harness wrapper in a child pwsh. Output streams
            # are redirected to files so they can be asserted on; the
            # process exit code is captured directly off the process
            # object to avoid `$LASTEXITCODE` races.
            function Invoke-LockedRestoreHarness {
                $script:stdoutFile = Join-Path $script:tempRoot 'stdout.txt'
                $script:stderrFile = Join-Path $script:tempRoot 'stderr.txt'
                $proc = Start-Process -FilePath pwsh -ArgumentList @(
                    '-NoProfile',
                    '-File', $script:harnessPath
                ) -NoNewWindow -Wait -PassThru `
                    -RedirectStandardOutput $script:stdoutFile `
                    -RedirectStandardError $script:stderrFile
                $script:runExitCode = $proc.ExitCode
                $script:runOutput = (Get-Content -LiteralPath $script:stdoutFile -Raw) + (Get-Content -LiteralPath $script:stderrFile -Raw)
            }
        }

        BeforeEach {
            Write-LockedRestoreHarness
        }

        AfterEach {
            if ($script:tempRoot -and (Test-Path -LiteralPath $script:tempRoot)) {
                Remove-Item -LiteralPath $script:tempRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
            if ($script:outsideRoot -and (Test-Path -LiteralPath $script:outsideRoot)) {
                Remove-Item -LiteralPath $script:outsideRoot -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        It 'scopes Remove-ObjDirectory to the repo root ($PSScriptRoot\..)' {
            # The copy lives in $script:tempRoot\scripts\, so the
            # script's $PSScriptRoot\.. is $script:tempRoot itself.
            # obj/ there must be deleted by Remove-ObjDirectory; obj/
            # outside the harness (in $script:outsideRoot) must
            # survive because the search is rooted at $script:tempRoot.
            Invoke-LockedRestoreHarness
            $script:runExitCode | Should -Be 0

            Test-Path -LiteralPath $script:objInside  | Should -BeFalse
            Test-Path -LiteralPath $script:objOutside | Should -BeTrue
        }

        It 'invokes four locked-mode restores; the two blocked ones carry the unreachable configfile' {
            Invoke-LockedRestoreHarness
            $script:runExitCode | Should -Be 0

            # Four Invoke-LockRestore calls in the recording log prove
            # four restore attempts; the `dotnet` shim lines carry the
            # exact argv so we can check --locked-mode and --no-cache.
            $log = Get-Content -LiteralPath $script:invocationLog
            $restoreCalls = @($log | Where-Object { $_ -match '^Invoke-LockRestore ' })
            $restoreCalls.Count | Should -Be 4

            $dotnetCalls = @($log | Where-Object { $_ -match '^dotnet ' })
            $dotnetCalls.Count | Should -Be 4
            foreach ($line in $dotnetCalls) {
                $line | Should -Match '--locked-mode'
                $line | Should -Match '--no-cache'
                $line | Should -Match 'GanttCreator\.slnx'
            }

            # Runs 3 (control) and 4 (offline) point at the generated
            # blocked NuGet.config and disable NuGetAudit (documented in
            # the script header).
            $blockedCalls = @($dotnetCalls | Where-Object { $_ -match '--configfile' })
            $blockedCalls.Count | Should -Be 2
            foreach ($line in $blockedCalls) {
                $line | Should -Match 'nuget-blocked-[0-9a-f-]+\.config'
                # The cmd/pwsh -File stub re-parses the attached
                # `-p:NuGetAudit=false` token into `-p NuGetAudit=false`;
                # both forms are equivalent MSBuild property syntax and the
                # real script passes the attached form.
                $line | Should -Match '-p[: ]NuGetAudit=false'
            }
        }

        It 'fails the negative control (blocked source, empty cache) and still exits 0' {
            # The stub `dotnet` exits non-zero exactly when NUGET_PACKAGES
            # points at the isolated empty directory, mirroring the real
            # blocked-source behaviour. An overall exit 0 therefore proves
            # the control ran with isolation, was expected to fail, did
            # fail, and did not abort the sequence before the offline run.
            Invoke-LockedRestoreHarness
            $script:runExitCode | Should -Be 0

            $log = Get-Content -LiteralPath $script:invocationLog
            $isolated = @($log | Where-Object { $_ -match '^dotnet .* nugget_packages=set' })
            $isolated.Count | Should -Be 1 -Because 'exactly the negative-control restore must run with the isolated empty NUGET_PACKAGES'
            $isolated[0] | Should -Match '--configfile'
        }

        It 'runs Remove-ObjDirectory before each success-expected restore with the control between runs 2 and 4' {
            # The recording wrappers append one line per helper call
            # in invocation order, so the helper-call lines in the
            # log ARE the execution trace. The `dotnet` shim also
            # appends a line per `dotnet` call; we filter to the
            # helper-call lines so this assertion is independent of
            # the dotnet-argv check above.
            Invoke-LockedRestoreHarness
            $script:runExitCode | Should -Be 0

            $log = Get-Content -LiteralPath $script:invocationLog
            $helperTrace = @($log | Where-Object { $_ -match '^(Remove-ObjDirectory|Invoke-LockRestore)' })
            # Labels are the script's own top-level -Label arguments (the
            # harness now replays the script's real sequence, not a
            # hardcoded abbreviated copy of it). The control does not need
            # a preceding obj/ cleanup, so no Remove-ObjDirectory line
            # precedes it.
            $helperTrace | Should -Be @(
                'Remove-ObjDirectory',
                'Invoke-LockRestore Run 1 -- generate lock files (no cache)',
                'Remove-ObjDirectory',
                'Invoke-LockRestore Run 2 -- honour lock files from clean state',
                'Invoke-LockRestore Run 3 control -- empty cache with blocked sources must fail',
                'Remove-ObjDirectory',
                'Invoke-LockRestore Run 4 -- locked restore with no reachable source (offline proof)'
            )
        }
    }
}
