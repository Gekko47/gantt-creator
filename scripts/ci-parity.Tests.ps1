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
        # Recursively finds all `dotnet test` command invocations in a
        # PowerShell AST. Using the AST (rather than regex) preserves quoted
        # separators, so a `;` or `|` inside a quoted string does not split
        # one invocation into two, and arguments with spaces are kept intact.
        # The input AST is produced by a single ParseInput call over the
        # complete script text, so backtick continuations are already
        # resolved and here-string prose and comments are naturally excluded:
        # a here-string is a string literal (never a CommandAst) and comments
        # belong to no AST node at all.
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

        # Returns one target-token count per `dotnet test` invocation, in
        # source order. Target tokens are literal project/solution paths
        # (.csproj/.vbproj/.fsproj/.slnx/.slnf/.sln) or a Solution/Project-named
        # PowerShell variable that the same script proves is a scalar string
        # (a [string]-constrained parameter or a [string] cast assignment).
        # Splatted variables, sub-expressions, and unverifiable variables are
        # rejected: dotnet test accepts exactly ONE positional project or
        # solution, and a collection or splat delivers many values (or none),
        # never one target. The healthy form is exactly 1 per invocation: an
        # aggregate total can hide a zero-target invocation alongside a
        # multi-target one, or flag two healthy invocations as one violation.
        function Get-DotnetTestProjectTokenCount {
            param([string]$ScriptText)
            $counts = @()
            # Parse the complete script text ONCE. The AST resolves line
            # continuations, quoted separators, here-strings, and comments
            # natively; the replaced per-line Join-LogicalLines pass parsed
            # each physical line separately, so here-string content was read
            # as executable code and counted as phantom invocations.
            $tokens = $null
            $parseErrors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$tokens, [ref]$parseErrors)
            # A missing AST or any parse error is a hard failure: unparseable
            # script text must not return an empty count list and silently
            # bypass the W14 validation loop. Returning an empty array here
            # would be indistinguishable from "zero dotnet test invocations",
            # so the caller's foreach would simply not iterate and the
            # exactly-one-target check would never fire. Throwing forces the
            # violation surface instead of masking it.
            if (-not $ast -or $parseErrors.Count -gt 0) {
                throw "Get-DotnetTestProjectTokenCount: script text failed to parse (parse errors: $($parseErrors.Count)). Unparseable input cannot be validated; treat as a W14 violation."
            }

            # A variable reference counts as a target only when the script
            # itself proves the variable is a scalar string: a [string]-typed
            # parameter (param([string]$Solution = ...)) or a [string] cast
            # assignment ([string]$Solution = ...). Anything else ($Projects
            # collections, splats, unresolved names) must not pass as one
            # positional target.
            #
            # The declaration is keyed by BOTH variable name and enclosing
            # lexical scope. A bare name-only key would treat every same-named
            # variable as one declaration: a [string]$Solution in function foo
            # would also vouch for a collection $Solution in function bar, or
            # for a script-scope $Projects collection named $Solution. Resolving
            # each command variable against its own enclosing scope before
            # applying scalar validation keeps same-named variables in
            # different functions or script scope independent.
            $scalarVars = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.HashSet[string]]]::new([System.StringComparer]::OrdinalIgnoreCase)
            function Add-ScalarVar {
                param([System.Management.Automation.Language.Ast]$Node, [string]$VarName)
                $scope = Get-EnclosingScopeName -Node $Node
                if (-not $scalarVars.ContainsKey($scope)) {
                    $scalarVars[$scope] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                }
                $null = $scalarVars[$scope].Add($VarName)
            }
            function Get-EnclosingScopeName {
                param([System.Management.Automation.Language.Ast]$Node)
                # The nearest FunctionDefinitionAst is the lexical scope; if
                # there is none the declaration lives in script scope. This
                # mirrors PowerShell's name resolution: a command inside a
                # function first looks at that function's parameters and
                # locals before falling back to script/global scope.
                $ancestor = $Node
                while ($ancestor) {
                    if ($ancestor -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                        return $ancestor.Name
                    }
                    $ancestor = $ancestor.Parent
                }
                return '__ScriptScope__'
            }
            function Get-EnclosingScopeChain {
                param([System.Management.Automation.Language.Ast]$Node)
                # Returns the ordered chain of enclosing scope names, from
                # the innermost lexical scope out to script scope. A variable
                # declared in any enclosing scope (a function's parameter, a
                # variable in a parent function, or a script-scope variable)
                # is visible to the command, so scalar validation must walk
                # the full chain rather than checking only the immediate
                # function name.
                $chain = [System.Collections.Generic.List[string]]::new()
                $ancestor = $Node
                while ($ancestor) {
                    if ($ancestor -is [System.Management.Automation.Language.FunctionDefinitionAst]) {
                        $chain.Add($ancestor.Name)
                    }
                    $ancestor = $ancestor.Parent
                }
                # Always include script scope as the outermost scope.
                $chain.Add('__ScriptScope__')
                return $chain
            }
            foreach ($typeConstraint in $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.TypeConstraintAst] }, $true)) {
                $typeName = $typeConstraint.TypeName.FullName -replace '^System\.', ''
                if ($typeName -ieq 'string') {
                    $parent = $typeConstraint.Parent
                    if ($parent -is [System.Management.Automation.Language.ParameterAst]) {
                        # param([string]$Solution = ...): the parameter's own
                        # scope is its enclosing function (or script scope for
                        # a script-level param block).
                        Add-ScalarVar -Node $parent -VarName $parent.Name.VariablePath.UserPath
                    } elseif ($parent -is [System.Management.Automation.Language.ConvertExpressionAst]) {
                        # For a [string] cast assignment ([string]$Solution = ...)
                        # the cast's operand variable is exposed as Child (the
                        # Expression property is null on an assignment LHS).
                        # Only record the operand variable when the cast is the
                        # LEFT side of an AssignmentStatementAst; a bare cast
                        # used as a value expression (e.g. [string]$x in a
                        # command argument) must not register its operand as a
                        # scalar variable.
                        if ($parent.Parent -is [System.Management.Automation.Language.AssignmentStatementAst]) {
                            $operand = $parent.Child
                            if ($operand -is [System.Management.Automation.Language.VariableExpressionAst]) {
                                Add-ScalarVar -Node $parent -VarName $operand.VariablePath.UserPath
                            }
                        }
                    }
                }
            }

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
                    if ($elem -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                        # `--diag` is parsed as a string constant (not a
                        # CommandParameterAst) in the `dotnet test` command
                        # line; the next element is its value, not a target.
                        if ($elem.Value -eq '--diag') {
                            $skipNext = $true
                        } elseif ($elem.Value -match '(?i)\.(csproj|vbproj|fsproj|slnx|slnf|sln)$') {
                            $count++
                        }
                        continue
                    }
                    if ($elem -is [System.Management.Automation.Language.VariableExpressionAst]) {
                        # In this PowerShell runtime, VariablePath.UserPath
                        # omits '$'; a splatted variable (@Splat) is never one
                        # positional target.
                        if (-not $elem.Splatted) {
                            $name = $elem.VariablePath.UserPath
                            if ($name -match '(?i)(Solution|Project)') {
                                # Resolve the variable against its enclosing scope chain
                                # before applying scalar validation, so a
                                # [string]$Solution declared in function foo does
                                # not vouch for a same-named collection in
                                # function bar or script scope. Walk the full
                                # chain (innermost function out to script
                                # scope) so a variable declared in any parent
                                # scope counts, not just the immediate scope.
                                $scopeChain = Get-EnclosingScopeChain -Node $elem
                                foreach ($scope in $scopeChain) {
                                    if ($scalarVars.ContainsKey($scope) -and $scalarVars[$scope].Contains($name)) {
                                        $count++
                                        break
                                    }
                                }
                            }
                        }
                        continue
                    }
                    # Any other expression type (sub-expressions, casts, nested
                    # commands) cannot be verified as a single positional
                    # project/solution target; it is never counted.
                }
                $counts += $count
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
        # Only a verified-scalar [string] Solution/Project variable counts as
        # a target, so the healthy fragment declares it the way the real
        # entry points do (test-non-office.ps1: param([string]$Solution = ...)).
        $healthy = @'
param([string]$Solution = 'GanttCreator.slnx', [string]$Configuration = 'Release')
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
param([string]$Solution = 'GanttCreator.slnx')
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

    It 'tripwire ignores dotnet test prose inside a here-string (positive control)' {
        # The counter parses the complete script text ONCE. A here-string is
        # a string literal, so its prose can never be a CommandAst; the
        # replaced per-line parser read that content as executable code and
        # counted a phantom invocation next to the real one.
        $prose = @"
param([string]`$Solution = 'GanttCreator.slnx')
`$notes = @'
dotnet test A.Tests/A.Tests.csproj -c Release
'@
dotnet test `$Solution -c Release --no-build --no-restore
"@
        @(Get-DotnetTestProjectTokenCount -ScriptText $prose) | Should -Be @(1)
    }

    It 'tripwire rejects a collection variable target that is not a verified scalar string' {
        # $Projects matched the old name-only regex and was misread as one
        # positional target. A variable the script does not prove to be a
        # scalar [string] must count as zero so the invocation is flagged.
        $collection = @'
dotnet test $Projects
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $collection) | Should -Be @(0)
    }

    It 'tripwire accepts a [string]-cast-assigned variable target (positive control)' {
        # [string]$Solution = ... proves the variable is a scalar string even
        # outside a param block; the verified variable counts as one target.
        $assigned = @'
[string]$Solution = 'GanttCreator.slnx'
dotnet test $Solution -c Release --no-build --no-restore
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $assigned) | Should -Be @(1)
    }

    It 'tripwire rejects a splatted variable target (pinning control)' {
        # A splat can deliver many or zero values; it is never one positional
        # project/solution target.
        @(Get-DotnetTestProjectTokenCount -ScriptText 'dotnet test @SolutionArgs') | Should -Be @(0)
    }

    It 'same-named variables in different lexical scopes remain independent (parity fixture)' {
        # The name-only scalar tracking defect treated every same-named
        # variable as one declaration: a [string]$Solution declared in one
        # function would vouch for a same-named collection in another, or for
        # a script-scope $Projects collection that happened to be named
        # $Solution. The fixture proves the scope-aware fix:
        #   - the function-scoped [string]$Solution counts as one target;
        #   - the same-named collection in a sibling function counts as zero;
        #   - the script-scope collection named $Solution counts as zero.
        $fixture = @'
function Use-Solution {
    param([string]$Solution = 'a.slnx')
    dotnet test $Solution
}
function Use-Collection {
    $Solution = @('x.csproj', 'y.csproj')
    dotnet test $Solution
}
$Solution = @('p.csproj', 'q.csproj')
dotnet test $Solution
'@
        @(Get-DotnetTestProjectTokenCount -ScriptText $fixture) | Should -Be @(1, 0, 0)
    }

    It 'tripwire throws on unparseable script text (positive control for parse-error hard failure)' {
        # The W14 validation loop iterates over whatever
        # Get-DotnetTestProjectTokenCount returns. If that function returned
        # an empty array on parse errors (instead of throwing), the caller's
        # foreach would simply not iterate and the exactly-one-target check
        # would silently pass. The parse-error branch now throws so that an
        # unparseable script is a visible W14 violation rather than a masked
        # empty result. This positive control proves the throw branch fires
        # when Parser.ParseInput produces parse errors, and guards against a
        # future edit that silently weakens the guard back to a soft return.
        #
        # `"dotnet test @"` is syntactically invalid: the `@` at statement
        # end is not a valid token start, so Parser.ParseInput reports a
        # parse error while still returning a non-null AST (the parser can
        # recover enough to produce a partial tree). Because the text still
        # contains `dotnet test` it reaches the counter via the caller's
        # regex filter, so the throw branch is the only way to surface the
        # unparseable input as a W14 violation.
        $unparseable = 'dotnet test @'
        { Get-DotnetTestProjectTokenCount -ScriptText $unparseable } | Should -Throw
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