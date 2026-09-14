#requires -Version 7
<#
.SYNOPSIS
    Pester tripwires for CI parity between .github/workflows/ci.yml and
    the scripts/ gate entry points.

.DESCRIPTION
    The workflow must not inline logic that already exists in scripts/,
    and version-sensitive tool pins (Pester / PSScriptAnalyzer /
    actionlint SHA-256) must come from scripts/tool-versions.psd1 as a
    single source of truth. The defect class prevented here is "local
    says PASS, CI says FAIL because the two views diverged":
      - W8: ci.yml inlined `Invoke-Pester -Script` (Pester 4 param, removed
        in Pester 5); the first CI run would have failed even though the
        local gate was green. Now: no inline Pester, no inline dotnet
        test/build/publish (all delegate to a scripts/*.ps1 entry point).
      - W11: actionlint SHA-256 and PSScriptAnalyzer version were each
        declared in two places (ci.yml + a local script); one was patched
        without the other. Now both pull from tool-versions.psd1.

    Exit 0 on clean; Pester reports failures to the runner.
#>

BeforeAll {
    $repoRoot = Split-Path -Parent $PSScriptRoot
    $ciPath = Join-Path $repoRoot '.github\workflows\ci.yml'
    $script:ciText = Get-Content -LiteralPath $ciPath -Raw

    # Extracts the body of one workflow step (from its `- name:` line to the
    # next `- name:` line) so assertions are scoped to a single step instead
    # of matching anywhere in the file.
    function Get-CiStepBlock {
        param([string]$Text, [string]$StepName)
        $lines = $Text -split "`r?`n"
        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*- name:\s*(.+?)\s*$') {
                if ($Matches[1] -eq $StepName) { $start = $i; break }
            }
        }
        if ($start -lt 0) { return $null }
        $body = New-Object System.Collections.Generic.List[string]
        for ($i = $start + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*- name:') { break }
            $null = $body.Add($lines[$i])
        }
        return ($body -join "`n")
    }
}

Describe 'ci.yml parity tripwires (W8)' {
    It 'has a Script lint (Pester) step' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script lint (Pester)'
        $block | Should -Not -BeNullOrEmpty
    }

    It 'script lint step delegates to scripts/test-scripts.ps1 (verify-quick parity)' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script lint (Pester)'
        $block | Should -Match 'test-scripts\.ps1'
    }

    It 'workflow never calls Invoke-Pester inline' {
        # Inline Pester invocation cannot stay in sync with the pinned
        # script gate; the parameter surface differs between Pester 5 and 6.
        # Strip full-line comments (matching PSSA tripwire behaviour) so the
        # forbidden-pattern assertion ignores comment-only mentions and still
        # detects Invoke-Pester in executable workflow content.
        $ciCode = ($script:ciText -split "`n") | Where-Object { $_ -notmatch '^\s*#' }
        ($ciCode -join "`n") | Should -Not -Match 'Invoke-Pester'
    }

    It 'format step delegates to dotnet format (allow-listed: native MSBuild command, not a script duplicate)' {
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Format check (production only; tests/ tolerated per tests/Directory.Build.props)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'dotnet format'
    }

    It 'workflow does not inline dotnet test/build/publish (must call scripts/*.ps1 or a vetted native command)' {
        # Verify scripts are the script-gate source of truth; the workflow
        # must call them, not re-implement the commands. The build, AddIn
        # publish, and non-Office test commands live in
        # scripts/build-release.ps1, scripts/publish-addin.ps1, and
        # scripts/test-non-office.ps1, shared with verify-quick.ps1 and
        # verify.ps1 (the W11/Phase C pairing completed here).
        # The tripwire is the full negative-match contract: the only
        # inline `run: dotnet <command>` lines allowed in ci.yml are the
        # vetted native MSBuild commands `dotnet format` and
        # `dotnet restore --locked-mode` (see docs/05-GIT-QUALITY.md
        # "CI parity (W8)"). Any other inline dotnet gate command -- a
        # `dotnet test` with or without trailing arguments, `dotnet
        # build`, `dotnet publish`, or a future gate command -- is a
        # violation, whether it sits on a single `run:` line, a YAML
        # sequence `- run: dotnet ...` line, or inside a multiline
        # `run: |` block. (?mi) makes ^ match each line start so the
        # tripwire sees the command at any indentation; `\s*`/`\s+`
        # consume extra whitespace, `-?\s*` accepts a YAML sequence
        # dash before the `run:` label, and command-line options are
        # irrelevant because the match stops at the command word.
        $script:ciText | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
    }

    It 'inline dotnet gate regex fires on deliberate violations and not on vetted commands (positive control)' {
        # Guards the negative tripwire above: if the regex cannot detect
        # a single-line `run: dotnet test ...` with trailing args, the
        # YAML sequence form `- run: dotnet test ...`, or an indented
        # `dotnet test` inside a multiline block, the anti-inline
        # assertion could never fire and an inline slip in ci.yml would
        # silently pass. The vetted commands must, conversely, not fire.
        $single = @'
steps:
  - name: Test
    run: dotnet test GanttCreator.slnx -c Release --filter 'Category!=OfficeIntegration'
'@
        $single | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        $dashed = @'
steps:
  - name: Test
    - run: dotnet test GanttCreator.slnx -c Release --no-build
'@
        $dashed | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        $block = @'
steps:
  - name: Test
    run: |
      Set-Location "$env:GITHUB_WORKSPACE"
      dotnet test GanttCreator.slnx --no-build
'@
        $block | Should -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'

        'run: dotnet restore --locked-mode' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        'run: dotnet format GanttCreator.slnx --verify-no-changes --exclude tests' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        '- run: dotnet restore --locked-mode' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
        '- run: dotnet format GanttCreator.slnx --verify-no-changes --exclude tests' | Should -Not -Match '(?mi)^\s*-?\s*(?:run:\s*)?dotnet\s+(?!format\b|restore\b)[a-z][a-z0-9-]*\b'
    }

    It 'test step delegates to scripts/test-non-office.ps1 (verify-quick parity)' {
        # The non-Office dotnet test command is version-sensitive and must
        # not be re-implemented inline in the workflow; the step points at
        # the same entry point verify-quick.ps1 step 12 runs.
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Test (OfficeIntegration excluded)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'test-non-office\.ps1'
    }

    It 'step extractor isolates exactly one step block (positive control)' {
        # Guards the helper itself: a block must stop at the next - name:
        # line, otherwise the delegation assertions could pass on a
        # neighbouring step's body.
        $fragment = @'
steps:
  - name: First
    run: echo one
  - name: Second
    run: echo two
'@
        $block = Get-CiStepBlock -Text $fragment -StepName 'First'
        $block | Should -Match 'echo one'
        $block | Should -Not -Match 'echo two'
    }
}

Describe 'dotnet test entry points pass exactly one project or solution (W14)' {
    # W14 (b54ccbd regression): scripts/test-non-office.ps1 passed five
    # .csproj paths to `dotnet test`, which accepts exactly ONE project or
    # solution argument; MSBuild failed with MSB1008 ("Only one project can
    # be specified") and every CI run went red even though the workflow
    # delegation itself was correct. The existing W8/W11/W13 tripwires only
    # police ci.yml content -- this Describe extends the parity net to the
    # scripts/*.ps1 entry points the workflow delegates to.

    BeforeAll {
        # Joins backtick line continuations so a logical command split over
        # several physical lines is scanned as one string.
        function Join-LogicalLines {
            param([string]$Text)
            $joined = New-Object System.Collections.Generic.List[string]
            $pending = ''
            foreach ($line in ($Text -split "`r?`n")) {
                if ($pending -ne '') {
                    $current = "$pending$line"
                } else {
                    $current = $line
                }
                if ($current -match '(?m)\x60$') {
                    $pending = $current -replace '(?m)\x60$', ''
                } else {
                    $null = $joined.Add($current)
                    $pending = ''
                }
            }
            if ($pending -ne '') { $null = $joined.Add($pending) }
            return $joined
        }

        # Returns one target-token count per `dotnet test` invocation, in
        # source order. Target tokens are file
        # references (.csproj/.vbproj/.fsproj/.slnx/.slnf/.sln) or
        # Solution/Project-named PowerShell variables; options, option
        # values, and filter literals are ignored. The healthy form is
        # exactly 1 per invocation: an aggregate total can hide a
        # zero-target invocation alongside a multi-target one, or flag two
        # healthy invocations as one violation.
        # Recursively finds all `dotnet test` command invocations in a
        # PowerShell AST. Using the AST (rather than regex) preserves quoted
        # separators, so a `;` or `|` inside a quoted string does not split
        # one invocation into two, and arguments with spaces are kept intact.
        function Find-DotnetTestCommands {
            param([System.Management.Automation.Language.Ast]$Ast, [ref]$Results)
            if (-not $Ast) { return }
            foreach ($cmd in $Ast.FindAll({ $true }, $true)) {
                if ($cmd -is [System.Management.Automation.Language.CommandAst] -and
                    $cmd.CommandElements -and
                    $cmd.CommandElements.Count -ge 2 -and
                    $cmd.CommandElements[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
                    $cmd.CommandElements[0].Value -ieq 'dotnet' -and
                    $cmd.CommandElements[1] -is [System.Management.Automation.Language.StringConstantExpressionAst] -and
                    $cmd.CommandElements[1].Value -ieq 'test') {
                    $Results.Value += $cmd
                }
            }
        }

        function Get-DotnetTestProjectTokenCount {
            param([string]$ScriptText)
            # Strip PowerShell block comments (<# ... #>) first: the synopsis
            # in test-non-office.ps1 mentions `dotnet test` in prose, which is
            # not an invocation and must not be counted as a zero-target
            # violation.
            $noBlock = $ScriptText -replace '(?s)<#.*?#>', ''
            $counts = @()
            foreach ($line in (Join-LogicalLines -Text $noBlock)) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                # Use the PowerShell AST to identify command invocations and
                # preserve quoted separators. The previous regex extracted one
                # tail from the complete logical line, which could not see
                # multiple dotnet test invocations on one line and split quoted
                # arguments on whitespace. Comments are naturally excluded by
                # the AST: they are not part of any CommandAst node.
                $ast = [System.Management.Automation.Language.Parser]::ParseInput($line, [ref]$null, [ref]$null)
                if (-not $ast) { continue }
                $commands = @()
                Find-DotnetTestCommands -Ast $ast -Results ([ref]$commands)
                foreach ($cmd in $commands) {
                    $count = 0
                    $skipNext = $false
                    $elements = $cmd.CommandElements
                    # Skip element 0 (the `test` subcommand); process the rest.
                    for ($i = 1; $i -lt $elements.Count; $i++) {
                        $elem = $elements[$i]
                        if ($skipNext) {
                            $skipNext = $false
                            continue
                        }
                        # Command parameters (options like -c, --diag) are not
                        # positional targets. --diag takes a diagnostic file path
                        # as its value; the next token is that value, not a
                        # project/solution target.
                        if ($elem -is [System.Management.Automation.Language.CommandParameterAst]) {
                            if ($elem.ParameterName -eq 'diag') {
                                $skipNext = $true
                            }
                            continue
                        }
                        # Get the token text for target matching. String constants
                        # (including quoted paths with spaces) and variables are
                        # the forms that can be targets; other expression types
                        # (splats, subexpressions, etc.) are not.
                        $token = $null
                        if ($elem -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                            $token = $elem.Value
                            # `--diag` is parsed as a string constant (not a
                            # CommandParameterAst) in the `dotnet test` command
                            # line; the next element is its value, not a target.
                            if ($token -eq '--diag') {
                                $skipNext = $true
                            }
                        } elseif ($elem -is [System.Management.Automation.Language.VariableExpressionAst]) {
                            # In this PowerShell runtime, VariableName omits '$';
                            # Extent.Text preserves the parsed variable token.
                            $token = $elem.Extent.Text
                        } else {
                            $token = $elem.Extent.Text
                        }
                        if ($token -match '(?i)\.(csproj|vbproj|fsproj|slnx|slnf|sln)$' -or
                            $token -match '(?i)^\$\w*(Solution|Project)\w*$') {
                            $count++
                        }
                    }
                    $counts += $count
                }
            }
            return $counts
        }
    }

    It 'every scripts/*.ps1 dotnet test invocation passes exactly one project or solution' {
        $violations = @()
        # $PSScriptRoot inside this Pester file resolves to scripts/, the
        # directory holding both the entry points and this test file.
        Get-ChildItem -Path $PSScriptRoot -Filter '*.ps1' |
            Where-Object { $_.Name -notlike '*Tests.ps1' } |
            ForEach-Object {
                $text = Get-Content -LiteralPath $_.FullName -Raw
                if ($text -match '(?i)dotnet\s+test(?:\s|$)') {
                    $index = 0
                    foreach ($n in @(Get-DotnetTestProjectTokenCount -ScriptText $text)) {
                        $index++
                        if ($n -ne 1) {
                            $violations += "$($_.Name) invocation ${index}: dotnet test has $n target token(s), expected exactly 1"
                        }
                    }
                }
            }
        $violations | Should -BeNullOrEmpty
    }

    It 'tripwire fires on the multiple-project defect form (positive control, b54ccbd regression)' {
        # The exact shape that shipped in b54ccbd and broke CI: several
        # backtick-continued .csproj paths on one logical dotnet test line.
        $broken = @'
dotnet test `
    tests/A.Tests/A.Tests.csproj `
    tests/B.Tests/B.Tests.csproj `
    tests/C.Tests/C.Tests.csproj `
    -c $Configuration --no-build --no-restore `
    --filter 'Category!=OfficeIntegration'
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $broken) | Should -Be @(3)
    }

    It 'tripwire accepts the single-solution form (negative control)' {
        $healthy = @'
dotnet test $Solution -c $Configuration --no-build --no-restore `
    --filter 'Category!=OfficeIntegration'
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $healthy) | Should -Be @(1)
    }

    It 'tripwire flags a dotnet test invocation with no project/solution target' {
        # Zero targets is equally ambiguous (repo-root default resolution);
        # the contract is exactly one, so 0 must also be detectable.
        $none = 'dotnet test -c Release --no-build --no-restore'
        @(Get-DotnetTestProjectTokenCount -ScriptText $none) | Should -Be @(0)

        # A bare `dotnet test` line (end-of-line, not whitespace, after the
        # command) must also enter the validation guard and be counted;
        # the exactly-one-target contract still flags it.
        @(Get-DotnetTestProjectTokenCount -ScriptText 'dotnet test') | Should -Be @(0)
    }

    It 'tripwire rejects a zero-target invocation alongside a valid one (positive control)' {
        # Aggregate validation would see 0 + 1 = 1 and pass; per-invocation
        # validation must flag the zero-target line independently.
        $mixed = @'
dotnet test -c Release --no-build --no-restore
dotnet test $Solution -c Release --no-build --no-restore
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $mixed) | Should -Be @(0, 1)
    }

    It 'tripwire accepts dotnet test --diag with a .sln-named diagnostic file and no target (positive control)' {
        # Guards the --diag option-value consumption: a diagnostic file path
        # following --diag must not be counted as a positional project/solution
        # target, and the invocation has no other target.
        $withDiag = @'
dotnet test --diag diag.sln -c Release --no-build --no-restore
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $withDiag) | Should -Be @(0)
    }

    It 'test-non-office.ps1 keeps the -Solution parameter (verify-quick parity)' {
        # The entry point must stay parameterised so verify-quick.ps1 and
        # ci.yml drive the same target; hardcoding project lists is what
        # enabled the W14 drift.
        $text = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'test-non-office.ps1') -Raw
        $text | Should -Match '\[string\]\$Solution\s*='
    }
}

Describe 'PSScriptAnalyzer gate delegates to Invoke-PssaGate (W13)' {
    It 'Script analyzer step routes through Invoke-PssaGate, not -EnableExit' {
        # W13: -EnableExit's function-level `exit` is swallowed by the
        # Tee/ForEach pipeline in Invoke-Step. The verify scripts route
        # through Invoke-PssaGate (scripts/verify-helpers.ps1); CI must do the
        # same or it re-introduces the W13 defect via a divergent path.
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Script analyzer (PSScriptAnalyzer)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'Invoke-PssaGate'
        # Strip comment lines (full-line # comments) before asserting the
        # gate is not invoked via -EnableExit: the step comment legitimately
        # names the forbidden switch as the reason for the delegation.
        $codeLines = ($block -split "`n") | Where-Object { $_ -notmatch '^\s*#' }
        ($codeLines -join "`n") | Should -Not -Match '-EnableExit'
    }
}

Describe 'tool-versions.psd1 single source of truth (W11)' {
    BeforeAll {
        $versionsPath = Join-Path $repoRoot 'scripts\tool-versions.psd1'
        $script:versions = Import-PowerShellDataFile -LiteralPath $versionsPath
        $script:lintCiText = Get-Content -LiteralPath (Join-Path $repoRoot 'scripts\lint-ci.ps1') -Raw
    }

    It 'pins actionlint version + SHA-256' {
        $script:versions.actionlint.Version | Should -Be '1.7.7'
        $script:versions.actionlint.Sha256 | Should -Match '^[0-9a-f]{64}$'
    }

    It 'actionlint step delegates to scripts/lint-ci.ps1 (verify-quick parity)' {
        # W11: the version + archive SHA-256 are single-sourced from
# scripts/tool-versions.psd1. The workflow delegates to
# scripts/lint-ci.ps1 which imports the psd1 and reads
# $versions.actionlint.Version / $versions.actionlint.Sha256
# internally, rather than carrying a literal mirror in the workflow.
# The anti-drift assertion is the delegation to lint-ci.ps1
# (mirroring the PSScriptAnalyzer test below), scoped to the
# actionlint step block so it cannot be satisfied by the PSScriptAnalyzer
# step's import.
        $block = Get-CiStepBlock -Text $script:ciText -StepName 'Workflow lint (actionlint)'
        $block | Should -Not -BeNullOrEmpty
        $block | Should -Match 'lint-ci\.ps1'
    }

    It 'actionlint SHA-256 in scripts/lint-ci.ps1 is sourced from tool-versions.psd1' {
        # W11: lint-ci.ps1 reads the pin from the psd1 at runtime; the
        # script source no longer carries the literal. The integrity
        # assertion follows the same indirection: both consumers read the
        # same property of the same psd1. $script:lintCiText is already
        # initialized by this Describe block's BeforeAll.
        $script:lintCiText | Should -Match 'Import-PowerShellDataFile'
        $script:lintCiText | Should -Match 'actionlint\.Sha256'
    }

    It 'ci.yml reads PSScriptAnalyzer version from scripts/tool-versions.psd1' {
        # W11: the version is consumed from the single source. The literal
        # "-RequiredVersion 1.25.0" is no longer in the workflow; what must
        # be true is that the workflow imports the psd1 and uses the
        # version variable.
        $script:ciText | Should -Match 'Import-PowerShellDataFile.*tool-versions\.psd1'
        $script:ciText | Should -Match 'PSScriptAnalyzer -RequiredVersion \$pssaVersion'
    }
}