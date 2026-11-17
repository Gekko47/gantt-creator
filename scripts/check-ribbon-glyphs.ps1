#requires -Version 7
<#
.SYNOPSIS
    Verifies that every ribbon glyph the shipped XML declares exists, in both palettes,
    at the pixel size the catalogue gives it, drawn in an approved stroke colour.

.DESCRIPTION
    R5.19 replaced every imageMso on the Gantt Creator tab with a themed Lucide glyph
    served by one getImage callback. That moves the failure mode: an imageMso that names
    nothing renders blank with no error, and so does a glyph asset that is missing, is the
    wrong size, or was rasterised in a colour nobody can see on the ribbon it lands on.
    This script is the gate for all three.

    It checks five things:

      1. no imageMso remains anywhere in the shipped ribbon;
      2. every control that declares getImage is in the generated catalogue, and every
         catalogued control is in the ribbon — so the two cannot drift;
      3. each catalogued glyph exists in both palettes at its declared size;
      4. each PNG is exactly that size and its ink is the approved stroke colour for that
         palette, measured as an alpha-weighted mean so antialiasing cannot fail the gate
         and a wrong colour cannot pass it;
      5. the six state glyphs exist in both palettes at 16 px.

    A glyph image that matches the naming grammar is an orphan if nothing references it,
    and that is a failure too: it means the catalogue and the assets disagree.

    The validator runs its own positive test first, on synthesised images, so a broken
    checker cannot report success. It does not need the git-ignored Office reference
    documents; those still gate the ids that remain in `check-ribbon-icons.ps1`.

.PARAMETER RibbonPath
    The shipped RibbonX document, relative to the repository root.

.PARAMETER CataloguePath
    The generated catalogues, relative to the repository root.

.PARAMETER ResourcePath
    The embedded PNG folder, relative to the repository root.

.EXAMPLE
    pwsh -NoProfile -File scripts/check-ribbon-glyphs.ps1
#>
[CmdletBinding()]
param(
    [string]$RibbonPath = 'src/GanttCreator.AddIn/RibbonResources/Ribbon.xml',
    [string]$CataloguePath = 'src/GanttCreator.AddIn/RibbonGlyphCatalog.g.cs',
    [string]$ResourcePath = 'src/GanttCreator.AddIn/Resources'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$repoRoot = Split-Path -Parent $PSScriptRoot
$channelTolerance = 12

# ADR-0063 D1: the two approved stroke colours. A light-palette glyph is drawn for a
# light ribbon and a dark-palette glyph for a dark one, so swapping them is a defect,
# not a preference.
$approvedStrokes = @{
    light = @(0x1F, 0x1F, 0x1F)
    dark  = @(0xF2, 0xF2, 0xF2)
}

function Get-GlyphMeasurement {
    [CmdletBinding()]
    param([string]$Path)

    $bitmap = [System.Drawing.Bitmap]::new($Path)
    try {
        $weight = 0.0
        $red = 0.0
        $green = 0.0
        $blue = 0.0

        for ($y = 0; $y -lt $bitmap.Height; $y++) {
            for ($x = 0; $x -lt $bitmap.Width; $x++) {
                $pixel = $bitmap.GetPixel($x, $y)
                if ($pixel.A -eq 0) {
                    continue
                }

                $weight += $pixel.A
                $red += $pixel.R * $pixel.A
                $green += $pixel.G * $pixel.A
                $blue += $pixel.B * $pixel.A
            }
        }

        if ($weight -eq 0) {
            return $null
        }

        return @{
            Width  = $bitmap.Width
            Height = $bitmap.Height
            Red    = $red / $weight
            Green  = $green / $weight
            Blue   = $blue / $weight
        }
    }
    finally {
        $bitmap.Dispose()
    }
}

function Test-GlyphFile {
    [CmdletBinding()]
    param(
        [string]$Path,
        [int]$Size,
        [string]$Palette
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return 'file is missing'
    }

    $measured = Get-GlyphMeasurement -Path $Path
    if ($null -eq $measured) {
        return 'file has no ink'
    }

    if ($measured.Width -ne $Size -or $measured.Height -ne $Size) {
        return ('is {0}x{1}, expected {2}x{2}' -f $measured.Width, $measured.Height, $Size)
    }

    $stroke = $approvedStrokes[$Palette]
    if ([Math]::Abs($measured.Red - $stroke[0]) -gt $channelTolerance -or
        [Math]::Abs($measured.Green - $stroke[1]) -gt $channelTolerance -or
        [Math]::Abs($measured.Blue - $stroke[2]) -gt $channelTolerance) {
        return ('stroke is ~#{0:X2}{1:X2}{2:X2}, expected #{3:X2}{4:X2}{5:X2}' -f
            [int]$measured.Red, [int]$measured.Green, [int]$measured.Blue,
            $stroke[0], $stroke[1], $stroke[2])
    }

    return $null
}

# --- Positive test ---------------------------------------------------------------------
# Synthesised images prove the checker fails what it must fail. Without this, a broken
# measurement would report every asset as fine.
$selfTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('gantt-glyph-selftest-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $selfTestRoot -Force
try {
    function Get-TestGlyph {
        [CmdletBinding()]
        param([string]$Name, [int]$Size, [System.Drawing.Color]$Colour)

        $path = Join-Path $selfTestRoot $Name
        $bitmap = [System.Drawing.Bitmap]::new($Size, $Size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.Clear([System.Drawing.Color]::Transparent)
                $pen = [System.Drawing.Pen]::new($Colour, 2.0)
                try {
                    $graphics.DrawLine($pen, 2, 2, $Size - 2, $Size - 2)
                }
                finally {
                    $pen.Dispose()
                }
            }
            finally {
                $graphics.Dispose()
            }

            $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally {
            $bitmap.Dispose()
        }

        return $path
    }

    $goodLight = Get-TestGlyph -Name 'good-light-16.png' -Size 16 -Colour ([System.Drawing.Color]::FromArgb(255, 0x1F, 0x1F, 0x1F))
    $goodDark = Get-TestGlyph -Name 'good-dark-16.png' -Size 16 -Colour ([System.Drawing.Color]::FromArgb(255, 0xF2, 0xF2, 0xF2))
    $wrongColour = Get-TestGlyph -Name 'wrong-colour-16.png' -Size 16 -Colour ([System.Drawing.Color]::FromArgb(255, 0xFF, 0x00, 0x00))
    $wrongSize = Get-TestGlyph -Name 'wrong-size-32.png' -Size 32 -Colour ([System.Drawing.Color]::FromArgb(255, 0x1F, 0x1F, 0x1F))

    $selfTestFailures = [System.Collections.Generic.List[string]]::new()

    if ($null -ne (Test-GlyphFile -Path $goodLight -Size 16 -Palette 'light')) {
        $selfTestFailures.Add('a correct light glyph was rejected')
    }

    if ($null -ne (Test-GlyphFile -Path $goodDark -Size 16 -Palette 'dark')) {
        $selfTestFailures.Add('a correct dark glyph was rejected')
    }

    if ($null -eq (Test-GlyphFile -Path $wrongColour -Size 16 -Palette 'light')) {
        $selfTestFailures.Add('a red glyph passed as the light stroke')
    }

    if ($null -eq (Test-GlyphFile -Path $wrongSize -Size 16 -Palette 'light')) {
        $selfTestFailures.Add('a 32 px glyph passed as a 16 px glyph')
    }

    if ($null -eq (Test-GlyphFile -Path (Join-Path $selfTestRoot 'absent.png') -Size 16 -Palette 'light')) {
        $selfTestFailures.Add('a missing file passed')
    }

    if ($null -eq (Test-GlyphFile -Path $goodLight -Size 16 -Palette 'dark')) {
        $selfTestFailures.Add('a light glyph passed as the dark stroke')
    }

    if ($selfTestFailures.Count -gt 0) {
        Write-Error ('check-ribbon-glyphs self-test failed: ' + ($selfTestFailures -join '; '))
        exit 1
    }
}
finally {
    Remove-Item -LiteralPath $selfTestRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# --- The shipped ribbon ----------------------------------------------------------------
$ribbonFile = Join-Path $repoRoot $RibbonPath
$catalogueFile = Join-Path $repoRoot $CataloguePath
$resourceDirectory = Join-Path $repoRoot $ResourcePath

foreach ($required in @($ribbonFile, $catalogueFile)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Error "Required file not found: $required"
        exit 1
    }
}

$failures = [System.Collections.Generic.List[string]]::new()

$ribbonText = Get-Content -LiteralPath $ribbonFile -Raw
$catalogueText = Get-Content -LiteralPath $catalogueFile -Raw

if ($ribbonText -match 'imageMso') {
    $failures.Add('the ribbon still names an imageMso; R5.19 removed every one of them')
}

# Controls served by the shared catalogue callback.
$declaredIds = [regex]::Matches($ribbonText, 'id="([^"]+)"[^>]*getImage="GetGlyph"') |
    ForEach-Object { $_.Groups[1].Value } |
    Sort-Object -Unique

# Controls whose picture follows their own state. Their assets are the state glyphs
# checked below, not catalogue entries.
$dynamicCallbacks = @(
    'GetAddEntityPrimaryImage'
    'GetToggleCellLockImage'
    'GetToggleChartLockImage'
    'GetColumnVisibilityImage'
)

foreach ($match in [regex]::Matches($ribbonText, 'id="([^"]+)"[^>]*getImage="([^"]+)"')) {
    $callback = $match.Groups[2].Value
    if ($callback -ne 'GetGlyph' -and $dynamicCallbacks -notcontains $callback) {
        $failures.Add(('{0} names an unknown image callback {1}' -f $match.Groups[1].Value, $callback))
    }
}

if ($declaredIds.Count -eq 0) {
    Write-Error 'No control in the ribbon declares getImage="GetGlyph".'
    exit 1
}

$catalogue = @{}
foreach ($match in [regex]::Matches($catalogueText, '\["([^"]+)"\] = new Entry\("([^"]+)", (\d+)\),')) {
    $catalogue[$match.Groups[1].Value] = @{
        Glyph = $match.Groups[2].Value
        Size  = [int]$match.Groups[3].Value
    }
}

if ($catalogue.Count -eq 0) {
    Write-Error "No control entries were parsed from $catalogueFile."
    exit 1
}

foreach ($id in $declaredIds) {
    if (-not $catalogue.ContainsKey($id)) {
        $failures.Add("$id declares getImage but is not in the glyph catalogue")
    }
}

foreach ($id in $catalogue.Keys) {
    if ($declaredIds -notcontains $id) {
        $failures.Add("$id is in the glyph catalogue but the ribbon does not use it")
    }
}

$checked = 0
$referenced = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

foreach ($id in $declaredIds) {
    if (-not $catalogue.ContainsKey($id)) {
        continue
    }

    $entry = $catalogue[$id]
    foreach ($palette in @('light', 'dark')) {
        $file = '{0}-{1}-{2}.png' -f $entry.Glyph, $palette, $entry.Size
        $null = $referenced.Add($file)

        $problem = Test-GlyphFile -Path (Join-Path $resourceDirectory $file) -Size $entry.Size -Palette $palette
        if ($null -ne $problem) {
            $failures.Add(("{0} -> {1}: {2}" -f $id, $file, $problem))
        }

        $checked++
    }
}

# State glyphs: their size is fixed by the control that shows them, not by the catalogue.
foreach ($match in [regex]::Matches($catalogueText, 'internal const string \w+ = "([^"]+)";')) {
    $slug = $match.Groups[1].Value
    foreach ($palette in @('light', 'dark')) {
        $file = '{0}-{1}-16.png' -f $slug, $palette
        $null = $referenced.Add($file)

        $problem = Test-GlyphFile -Path (Join-Path $resourceDirectory $file) -Size 16 -Palette $palette
        if ($null -ne $problem) {
            $failures.Add(("{0} -> {1}: {2}" -f $slug, $file, $problem))
        }

        $checked++
    }
}

# The eight glyphs the three add primaries can show, at 32 px. The set mirrors the
# generator's state table and is pinned by RibbonGlyphAssetsTests, which asserts every
# GanttEntityType maps into exactly these eight. The three concept slugs were added by the
# 2026-11-17 surface revision so a primary whose label names Delay Event, a procurement
# period or a Custom row cannot keep showing the family picture.
$primaryGlyphs = @(
    'add-activity'
    'add-milestone'
    'circle-dashed'
    'delineator'
    'spacer'
    'splitter'
    'procurement'
    'span-delay'
)
foreach ($slug in $primaryGlyphs) {
    foreach ($palette in @('light', 'dark')) {
        $file = '{0}-{1}-32.png' -f $slug, $palette
        $null = $referenced.Add($file)

        $problem = Test-GlyphFile -Path (Join-Path $resourceDirectory $file) -Size 32 -Palette $palette
        if ($null -ne $problem) {
            $failures.Add(("{0} -> {1}: {2}" -f $slug, $file, $problem))
        }

        $checked++
    }
}

# Orphans: a themed asset nothing references means the catalogue and the assets disagree.
$orphans = Get-ChildItem -LiteralPath $resourceDirectory -Filter '*.png' |
    Where-Object { $_.Name -match '-(light|dark)-(16|32)\.png$' } |
    Where-Object { -not $referenced.Contains($_.Name) } |
    Select-Object -ExpandProperty Name

foreach ($orphan in $orphans) {
    $failures.Add("$orphan is committed but nothing references it")
}

if ($failures.Count -gt 0) {
    Write-Error ("check-ribbon-glyphs: {0} violation(s):`n  - {1}" -f $failures.Count, ($failures -join "`n  - "))
    exit 1
}

Write-Host ('OK: {0} control(s) and {1} glyph asset(s) verified across both palettes, with no imageMso and no orphan.' -f $declaredIds.Count, $checked)
