# Single source of truth for external tool versions and integrity
# digests used by both CI (.github/workflows/ci.yml) and the local
# scripts/ entry points. The CI-parity and tool-version anti-drift
# Pester tests assert that ci.yml and the local scripts consume values
# from this file (or a string-identical copy), so version drift between
# local and CI cannot recur. See AGENTS.md "Test layers" and
# docs/04-TEST-STRATEGY.md.
@{
    # Pester. CI installs the exact Version below, never "newest >=
    # MinimumMajor": an open-ended install ran a different Pester on every
    # hosted runner and produced the intermittent parallel worker-startup
    # failure in CI runs 45193295738 / 45376223093. MinimumMajor stays the
    # floor the runner (scripts/test-scripts.ps1) enforces at runtime, and
    # an exact Version equal to it keeps local and CI identical.
    Pester = @{
        Version      = '6.2.0'
        MinimumMajor = 6
    }

    # PSScriptAnalyzer. 1.25.0 is the version CI installs and the version
    # the local .NET host has. Bumping either side without bumping the
    # other is caught by lint-ci.Tests.ps1 and pssa-gate.Tests.ps1.
    PSScriptAnalyzer = @{
        Version = '1.25.0'
    }

    # actionlint. Both the CI workflow and scripts/lint-ci.ps1 verify the
    # SHA-256 below before using the downloaded binary. Sha256 pins the
    # published actionlint 1.7.12 windows/amd64 zip; ExeSha256 pins the
    # actionlint.exe inside that archive (computed from a fresh download
    # whose archive hash matched the pin) so scripts/lint-ci.ps1 can also
    # re-verify an already-cached executable before every execution.
    actionlint = @{
        Version     = '1.7.12'
        Sha256      = '6e7241b51e6817ea6a047693d8e6fed13b31819c9a0dd6c5a726e1592d22f6e9'
        ExeSha256   = '54ca21be3de4c7cfa26914aa8b61bd76bf573ef3caac5f80d110558cdf241718'
        DownloadUrl = 'https://github.com/rhysd/actionlint/releases/download/v{0}/actionlint_{0}_windows_amd64.zip'
    }

    # Stryker.NET - the R3.13 mutation-testing tool, installed as a GLOBAL
    # `dotnet tool` by scripts/mutation-gate.ps1 (never a csproj
    # PackageReference: it is test-only and must not reach a production
    # project). Pinned here so local and CI cannot diverge (W11).
    #
    # 5.1.0 is the current latest stable (R3.13's first probe ran 5.0.0; the
    # diff-gate flags below were re-verified against 5.1.0 on 2026-11-17).
    # VERIFIED against the installed CLI, not from memory (evidence ledger in
    # docs/work-items/R3.13-mutation-testing.md):
    #   * `--since[:<committish>]` exists and enables diff-compare mode, so the
    #     changed-code scope D2 assumes is real rather than assumed.
    #   * `--threshold-high` / `--threshold-low` / `--break-at` exist, and
    #     `--threshold-low` is a hard gate returning non-zero.
    #   * `--project` / `--test-project` take explicit csproj paths, so the gate
    #     does NOT depend on Stryker parsing the .slnx solution file - the
    #     guide's ".slnx gaps" stop condition does not apply.
    'dotnet-stryker' = @{
        Version       = '5.1.0'
        ToolId        = 'dotnet-stryker'
        # The changed-code mutation threshold from docs/04-TEST-STRATEGY.md
        # "Coverage policy" (80%). Scoped to changed Core code, not the whole
        # project, so this is the score the gate judges.
        ThresholdHigh = 80
    }

    # ReportGenerator. Merges the per-test-project Cobertura files that
    # `dotnet test --collect:'XPlat Code Coverage'` writes into ONE report.
    # Without the merge no per-file number can be compared with the coverage
    # thresholds: each file holds only the classes its own test project
    # executed, so the Core.Tests file alone measured GanttCreator.Core at
    # 0.28% line while the union of every collected file is ~97%. Installed as
    # a GLOBAL `dotnet tool` by scripts/coverage-report.ps1 (never a csproj
    # PackageReference: it is a dev tool and must not reach a production
    # project). Pinned here so local and CI cannot diverge (W11).
    #
    # 5.6.0 VERIFIED against the installed CLI, not from memory:
    # `dotnet tool update --global dotnet-reportgenerator-globaltool`
    # reported "version '5.6.0'" and `dotnet tool list --global` agrees.
    'dotnet-reportgenerator-globaltool' = @{
        Version = '5.6.0'
        ToolId  = 'dotnet-reportgenerator-globaltool'
    }
}