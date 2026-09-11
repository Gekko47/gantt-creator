# Single source of truth for external tool versions and integrity
# digests used by both CI (.github/workflows/ci.yml) and the local
# scripts/ entry points. The CI-parity and tool-version anti-drift
# Pester tests assert that ci.yml and the local scripts consume values
# from this file (or a string-identical copy), so version drift between
# local and CI cannot recur. See AGENTS.md "Test layers" and
# docs/04-TEST-STRATEGY.md.
@{
    # Pester. CI installs newest >= 6.0; local runs 6.1.0. When a
    # known-broken version ships, add a minimum here and assert it from
    # Pester.
    Pester = @{
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
    # published actionlint 1.7.7 windows/amd64 zip; ExeSha256 pins the
    # actionlint.exe inside that archive (computed from a fresh download
    # whose archive hash matched the pin) so scripts/lint-ci.ps1 can also
    # re-verify an already-cached executable before every execution.
    actionlint = @{
        Version     = '1.7.7'
        Sha256      = '7f12f1801bca3d480d67aaf7774f4c2a6359a3ca8eebe382c95c10c9704aa731'
        ExeSha256   = '6d470a52039a433bccdf57bd65170885a5b23e511155d0de587c2c3ce0eb1c24'
        DownloadUrl = 'https://github.com/rhysd/actionlint/releases/download/v{0}/actionlint_{0}_windows_amd64.zip'
    }
}