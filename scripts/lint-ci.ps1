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

# Collect every workflow file under .github/workflows so the lint gate
# validates the whole set, not only ci.yml. Actionlint accepts multiple
# file arguments, so we pass them all in one invocation.
$workflowFiles = Get-ChildItem -LiteralPath $workflowDir -File -Filter '*.yml' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName
$workflowFiles += Get-ChildItem -LiteralPath $workflowDir -File -Filter '*.yaml' -ErrorAction SilentlyContinue |
    Select-Object -ExpandProperty FullName
$workflowFiles = $workflowFiles | Sort-Object -Unique

if ($workflowFiles.Count -eq 0) {
    Write-Error "No workflow files found under $workflowDir"
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
$expectedExeHash = $script:versions.actionlint.ExeSha256
$toolsDir = Join-Path $env:TEMP "actionlint-$version"
$exePath = Join-Path $toolsDir $exeName
$hashScript = Join-Path $PSScriptRoot 'actionlint-hash.ps1'

# Cached-executable integrity: an already-downloaded exe is never trusted
# blindly. Its SHA-256 is validated against the pinned ExeSha256 before
# every execution; a missing or tampered cache is removed and replaced by
# a fresh, hash-verified download instead of being used.
$exeVerified = $false
if (Test-Path -LiteralPath $exePath) {
    & $hashScript -FilePath $exePath -ExpectedHash $expectedExeHash
    if ($LASTEXITCODE -eq 0) {
        $exeVerified = $true
    }
    else {
        Write-Host 'Cached actionlint executable failed the SHA-256 check; removing the cache and redownloading...'
        Remove-Item -LiteralPath $toolsDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Download if not present (or if the cached copy just failed verification).
if (-not $exeVerified) {
    Write-Host "Downloading actionlint v$version..."
    try {
        if (-not (Test-Path -LiteralPath $toolsDir)) { New-Item -ItemType Directory -Path $toolsDir -Force | Out-Null }
        $zipPath = Join-Path $toolsDir $zipName
        # Explicit timeout: no script-wide convention exists, so a fixed
        # 300s ceiling is used. A stalled download must fail promptly rather
        # than hang the lint gate indefinitely on a dead connection.
        Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath -UseBasicParsing -TimeoutSec 300
        # Integrity check: the same SHA-256 pin the CI workflow enforces.
        # The pin is single-sourced from scripts/tool-versions.psd1 so the
        # two invocations can never disagree; ci-parity.Tests.ps1 enforces
        # the equality. Both pins must stay identical. The check itself lives
        # in scripts/actionlint-hash.ps1 so it can be unit-tested offline.
        & $hashScript -ZipPath $zipPath -ExpectedHash $expectedHash
        if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
        Expand-Archive -Path $zipPath -DestinationPath $toolsDir -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "Failed to download actionlint: $_"
        exit 1
    }

    # Verify the freshly extracted executable against its own pin before
    # use. A mismatch after a hash-verified archive download is a
    # supply-chain failure: fail hard rather than execute the binary.
    & $hashScript -FilePath $exePath -ExpectedHash $expectedExeHash
    if ($LASTEXITCODE -ne 0) {
        Remove-Item -LiteralPath $toolsDir -Recurse -Force -ErrorAction SilentlyContinue
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $exePath)) {
    Write-Error "actionlint executable not found after download: $exePath"
    exit 1
}

Write-Host "Running actionlint v$version against $($workflowFiles.Count) workflow file(s)..."
& $exePath -color @workflowFiles 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Error "actionlint found issues in $($workflowFiles -join ', ')"
    exit $LASTEXITCODE
}

Write-Host "lint-ci: PASS ($($workflowFiles.Count) workflow file(s) validated by actionlint v$version)"
exit 0