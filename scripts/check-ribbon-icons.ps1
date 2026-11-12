#requires -Version 7
# Validates that every imageMso in the shipped Ribbon.xml names a real Office
# icon. imageMso fails silently: an unknown id renders no glyph, with no error
# and no log, which is how the Data group once shipped nine identical table
# icons and three buttons with no icon at all (R5.16a).
#
# Two reference bases, both required in practice:
#   1. the Office 2010 Icons Gallery (.microsoft/Office2010IconsGallery.docx),
#      whose embedded customUI14.xml lists every gallery id;
#   2. the idMso table in [MS-CUSTOMUI]-250218.docx, where the id must appear
#      as a whole line in the document text.
# Both references are git-ignored, so the check reports "Skipped" and exits 0
# when they are absent. It never passes silently.
#
# The validator's own positive test runs first: a fabricated id must be
# rejected and a known id accepted, so a broken validator cannot report success.
# Exits 0 when every id passes; exits 1 with the offending ids otherwise.
[CmdletBinding()]
param(
    [string]$RibbonPath = 'src/GanttCreator.AddIn/RibbonResources/Ribbon.xml',
    [string]$GalleryPath = '.microsoft/Office2010IconsGallery.docx',
    [string]$CustomUiPath = '.microsoft/[MS-CUSTOMUI]-250218.docx'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Get-RibbonImageMsoIds {
    [CmdletBinding()]
    param([string]$Xml)

    [regex]::Matches($Xml, 'imageMso="([^"]+)"') |
        ForEach-Object { $_.Groups[1].Value } |
        Sort-Object -Unique
}

function Get-DocxEntryText {
    [CmdletBinding()]
    param([string]$Path, [string]$EntryPattern)

    $resolved = (Resolve-Path -LiteralPath $Path).ProviderPath
    $zip = [System.IO.Compression.ZipFile]::OpenRead($resolved)
    try {
        $entry = $zip.Entries |
            Where-Object { $_.FullName -match $EntryPattern } |
            Select-Object -First 1
        if ($null -eq $entry) {
            throw "Entry $EntryPattern not found in $Path"
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Test-ImageMsoIds {
    [CmdletBinding()]
    param(
        [string[]]$Ids,
        [string[]]$GalleryIds,
        [string[]]$CustomUiLines
    )

    $failures = New-Object System.Collections.Generic.List[string]
    foreach ($id in $Ids) {
        if ($GalleryIds -notcontains $id) {
            $failures.Add("$id : absent from the Office 2010 Icons Gallery")
        }

        if ($CustomUiLines -notcontains $id) {
            $failures.Add("$id : absent from the [MS-CUSTOMUI] idMso table")
        }
    }

    return $failures
}

# Positive test: the validator must reject a fabricated id and accept a real one.
# The fabricated id fails BOTH bases, so exactly two failures are expected.
$selfTestFailures = @(Test-ImageMsoIds -Ids @('DefinitelyNotARealImageMsoId') -GalleryIds @('Help') -CustomUiLines @('Help'))
if ($selfTestFailures.Count -ne 2) {
    Write-Error 'check-ribbon-icons self-test failed: a fabricated imageMso id was not rejected by both bases.'
    exit 1
}

$selfTestAccepted = @(Test-ImageMsoIds -Ids @('Help') -GalleryIds @('Help') -CustomUiLines @('Help'))
if ($selfTestAccepted.Count -ne 0) {
    Write-Error 'check-ribbon-icons self-test failed: a known imageMso id was rejected.'
    exit 1
}

$ribbonFile = Join-Path $repoRoot $RibbonPath
if (-not (Test-Path -LiteralPath $ribbonFile)) {
    Write-Error "Ribbon XML not found: $ribbonFile"
    exit 1
}

$imageMsoIds = @(Get-RibbonImageMsoIds -Xml (Get-Content -LiteralPath $ribbonFile -Raw))
if ($imageMsoIds.Count -eq 0) {
    Write-Error "No imageMso attributes found in $ribbonFile"
    exit 1
}

$galleryFile = Join-Path $repoRoot $GalleryPath
$customUiFile = Join-Path $repoRoot $CustomUiPath
if (-not (Test-Path -LiteralPath $galleryFile) -or -not (Test-Path -LiteralPath $customUiFile)) {
    Write-Host "Skipped: reference absent ($GalleryPath or $CustomUiPath); check $($imageMsoIds.Count) ids on a machine with .microsoft/."
    exit 0
}

$galleryXml = Get-DocxEntryText -Path $galleryFile -EntryPattern 'customUI14\.xml$'
$galleryIds = @(Get-RibbonImageMsoIds -Xml $galleryXml)

$customUiXml = Get-DocxEntryText -Path $customUiFile -EntryPattern 'word/document\.xml$'
$plainText = [regex]::Replace($customUiXml, '</w:p>', "`n")
$plainText = [regex]::Replace($plainText, '<[^>]+>', '')
$customUiLines = @($plainText -split "`n" | ForEach-Object { $_.Trim() })

$failures = Test-ImageMsoIds -Ids $imageMsoIds -GalleryIds $galleryIds -CustomUiLines $customUiLines
if ($failures.Count -gt 0) {
    Write-Error ("Unverified imageMso ids:" + [Environment]::NewLine + ($failures -join [Environment]::NewLine))
    exit 1
}

Write-Host "OK: $($imageMsoIds.Count) imageMso ids verified against the gallery and the idMso table."
exit 0
