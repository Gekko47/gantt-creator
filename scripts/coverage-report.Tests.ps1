#requires -Version 7
<#
    Pester tests for coverage-report.ps1. The merge tool version must come
    from scripts/tool-versions.psd1 (single source of truth, W11), and the
    no-silent-pass rules (no reports, tool failure, no packages) must be
    present in the script, because a merge that quietly reports a number the
    suite never produced is exactly the defect class this gate exists to
    prevent.
#>

BeforeAll {
    $script:repoRoot = Split-Path -Parent $PSScriptRoot
    $script:scriptPath = Join-Path $script:repoRoot 'scripts\coverage-report.ps1'
    $script:versionsPath = Join-Path $script:repoRoot 'scripts\tool-versions.psd1'
    $script:versions = Import-PowerShellDataFile -LiteralPath $script:versionsPath
    $script:scriptText = Get-Content -LiteralPath $script:scriptPath -Raw

    $script:fixture = [xml]@'
<coverage line-rate="0.9" branch-rate="0.8">
  <packages>
    <package name="GanttCreator.Core" line-rate="0.9693" branch-rate="0.8755" />
    <package name="GanttCreator.Core.Tests" line-rate="1" branch-rate="1" />
  </packages>
</coverage>
'@
}

Describe 'coverage-report.ps1 (pinned merge tool)' {
    Context 'tool pinning (W11)' {
        It 'pins dotnet-reportgenerator-globaltool in tool-versions.psd1' {
            $script:versions.'dotnet-reportgenerator-globaltool' | Should -Not -BeNullOrEmpty
            $script:versions.'dotnet-reportgenerator-globaltool'.ToolId |
                Should -BeExactly 'dotnet-reportgenerator-globaltool'
            $script:versions.'dotnet-reportgenerator-globaltool'.Version |
                Should -Match '^\d+\.\d+\.\d+$'
        }

        It 'rejects a missing pin' {
            # Positive test for the validator: feed it the computed-null
            # settings object and assert the error path fires.
            . $script:scriptPath
            { Assert-ReportGeneratorPin $null } |
                Should -Throw '*has no dotnet-reportgenerator-globaltool pin*'
        }

        It 'accepts the pinned settings object' {
            . $script:scriptPath
            { Assert-ReportGeneratorPin $script:versions.'dotnet-reportgenerator-globaltool' } |
                Should -Not -Throw
        }

        It 'calls the pin guard from the main body' {
            # A guard with no call site proves nothing.
            $script:scriptText | Should -Match 'Assert-ReportGeneratorPin \$settings'
        }
    }

    Context 'merge result parsing' {
        It 'reads a package line and branch rate from the merged report' {
            . $script:scriptPath
            $result = Get-PackageCoverageRate -Report $script:fixture -PackageName 'GanttCreator.Core'
            $result.Package | Should -BeExactly 'GanttCreator.Core'
            $result.LineRate | Should -Be 0.9693
            $result.BranchRate | Should -Be 0.8755
        }

        It 'returns null for a package the report does not contain' {
            . $script:scriptPath
            Get-PackageCoverageRate -Report $script:fixture -PackageName 'GanttCreator.Raster' |
                Should -BeNullOrEmpty
        }

        It 'excludes test assemblies from the production package list' {
            . $script:scriptPath
            $packages = @(Get-MergedPackageCoverage -Report $script:fixture)
            $packages.Count | Should -Be 1
            $packages[0].Package | Should -BeExactly 'GanttCreator.Core'
        }
    }

    Context 'no-silent-pass' {
        It 'refuses to merge when no report was collected' {
            $script:scriptText | Should -Match 'No coverage\.cobertura\.xml'
        }

        It 'fails when reportgenerator exits non-zero' {
            $script:scriptText | Should -Match 'reportgenerator failed with exit code'
        }

        It 'fails when the merged report has no packages' {
            $script:scriptText | Should -Match 'contains no packages'
        }

        It 'excludes generated compiler sources from the merged count' {
            # RegexGenerator output is machine-written; counting it as
            # uncovered Core code measures the compiler, not this repository.
            $script:scriptText | Should -Match 'filefilters:-\*\*/obj/\*\*'
        }
    }
}
