#requires -Version 7
<#
.SYNOPSIS
    Workflow lint gate (L6). Invokes pinned actionlint against .github/workflows/ci.yml.
    Wired into verify-quick.ps1 and CI.

.DESCRIPTION
    actionlint validates GitHub Actions workflow syntax, required fields,
    and common misconfigurations. We use a pinned binary (not npm) so the
    check runs on the windows-latest runner without extra tooling.

    Exit 0 on clean; exit 1 with annotated failures.

.NOTES
    Pinned version: 1.7.7 (latest stable at time of R0.8).
    Download URL pattern: https://github.com/rhysd/actionlint/releases/download/v{VERSION}/actionlint_{VERSION}_windows_amd64.zip
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$versionsPath = Join-Path $repoRoot 'scripts\tool-versions.psd1'
$workflowDir = Join-Path $repoRoot '.github\workflows'
$ciYml = Join-Path $workflowDir 'ci.yml'

if (-not (Test-Path -LiteralPath $ciYml)) {
    Write-Error "Workflow file not found: $ciYml"
    exit 1
}

# Tool versions + integrity hash are single-sourced from scripts/tool-versions.psd1
# (W11). ci-parity.Tests.ps1 asserts the same hash lives in ci.yml and here, so
# neither side can drift without the tripwire firing.
$script:versions = Import-PowerShellDataFile -LiteralPath $versionsPath
$version = $script:versions.actionlint.Version
$toolName = 'actionlint'
$exeName = "$toolName.exe"
$zipName = "actionlint_${version}_windows_amd64.zip"
$downloadUrl = $script:versions.actionlint.DownloadUrl -f $version
$expectedHash = $script:versions.actionlint.Sha256
$toolsDir = Join-Path $env:TEMP 'actionlint'
$exePath = Join-Path $toolsDir $exeName

# Download if not present
if (-not (Test-Path -LiteralPath $exePath)) {
    Write-Host "Downloading actionlint v$version..."
    try {
        if (-not (Test-Path -LiteralPath $toolsDir)) { New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null }
        $zipPath = Join-Path $toolsDir $zipName
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing
        # Integrity check: the same SHA-256 pin the CI workflow enforces.
        # The pin is single-sourced from scripts/tool-versions.psd1 so the
        # two invocations can never disagree; ci-parity.Tests.ps1 enforces
        # the equality. Both pins must stay identical. The check itself lives
        # in scripts/actionlint-hash.ps1 so it can be unit-tested offline.
        & (Join-Path $PSScriptRoot 'actionlint-hash.ps1') -ZipPath $zipPath -ExpectedHash $expectedHash
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
        Expand-Archive -Path $zipPath -DestinationPath $toolsDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "Failed to download actionlint: $_"
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $exePath)) {
    Write-Error "actionlint executable not found after download: $exePath"
    exit 1
}

Write-Host "Running actionlint v$version against $ciYml..."
& $exePath -color $ciYml 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "actionlint found issues in $ciYml"
    exit $LASTEXITCODE
}

Write-Host "lint-ci: PASS (ci.yml validated by actionlint v$version)"
exit 0