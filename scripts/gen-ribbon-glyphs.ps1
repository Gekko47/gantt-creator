#requires -Version 7
<#
.SYNOPSIS
    Regenerates the ribbon's committed glyph images from Lucide geometry.

.DESCRIPTION
    The RibbonX surface cannot be given an SVG and the add-in deliberately carries
    no SVG renderer, so every glyph is rasterised HERE, once, under review, and
    committed as a PNG. The runtime only decodes an embedded PNG.

    Two families of geometry are used:

      * the static control glyphs, one per control, at 32 px on a large control and
        16 px on a normal control or a menu row;
      * the state glyphs, where the picture follows a toggle or a row's own state:
        the column tick and dash, and the two padlock pairs.

    Every glyph is rendered twice, in a light stroke and a dark stroke, so the
    callback that answers Excel can pick the pair that matches the Office theme
    (ADR-0063).

    Geometry is Lucide's, verbatim. Each entry below is the `d` attribute Lucide
    publishes for that icon on its 24 x 24 view box, with stroke width 2 and round
    caps and joins; `rect:` and `circle:` entries are the SVG primitives Lucide
    uses in the same icons. Nothing here is retyped from a screenshot.

    WPF's geometry mini-language parses that `d` syntax directly, so the script
    carries no path parser of its own. That is also why rotated arcs such as the
    ones in `table-rows-split` render correctly without special handling.

.PARAMETER ResourceDirectory
    Where the PNGs are written. Defaults to the add-in's Resources folder.

.EXAMPLE
    pwsh -NoProfile -File scripts/gen-ribbon-glyphs.ps1

.NOTES
    Deterministic on a machine: the same script, runtime and assembly produce the
    same pixels, so a diff means the geometry or the palette changed. Run it, look
    at the PNGs, commit them.

    This script supersedes `scripts/gen-column-visibility-icons.ps1`, which drew
    the same two families by hand with a fixed black pen and therefore could not
    answer a dark ribbon.
#>
[CmdletBinding()]
param(
    [string]$ResourceDirectory = (Join-Path $PSScriptRoot '..\src\GanttCreator.AddIn\Resources')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

$viewBox = 24.0
$strokeWidth = 2.0

# ADR-0063: neutral strokes only, so a glyph never competes with an Office accent
# or a chart token. The light stroke clears 13:1 on the light ribbon; the dark
# stroke clears 11:1 on Dark Gray and Black.
$palettes = [ordered]@{
    light = [System.Windows.Media.Color]::FromRgb(0x1F, 0x1F, 0x1F)
    dark  = [System.Windows.Media.Color]::FromRgb(0xF2, 0xF2, 0xF2)
}

# Lucide geometry. A string is either a path `d`, `rect:x,y,w,h,rx`, or
# `circle:cx,cy,r` / `circle:cx,cy,r,fill` for Lucide's filled dots.
$glyphs = [ordered]@{
    'initialise-sheet'  = @(
        'rect:3,3,18,18,2'
        'M9 8h7'
        'M8 12h6'
        'M11 16h5'
    )
    'repair'            = @(
        'M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.106-3.105c.32-.322.863-.22.983.218a6 6 0 0 1-8.259 7.057l-7.91 7.91a1 1 0 0 1-2.999-3l7.91-7.91a6 6 0 0 1 7.057-8.259c.438.12.54.662.219.984z'
    )
    'validate'          = @(
        'M13 5h8'
        'M13 12h8'
        'M13 19h8'
        'm3 17 2 2 4-4'
        'm3 7 2 2 4-4'
    )
    'add-activity'      = @(
        'M6 5h12'
        'M4 12h10'
        'M12 19h8'
    )
    'add-milestone'     = @(
        'M12 13v8'
        'M12 3v3'
        'M4 6a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h13a2 2 0 0 0 1.152-.365l3.424-2.317a1 1 0 0 0 0-1.635l-3.424-2.318A2 2 0 0 0 17 6z'
    )
    # R5.16a surface revision 2026-11-17. Three concept rows in the add-entity menus
    # needed a picture that differs at 16 px, where detail is lost and only a
    # silhouette survives. `procurement` doubles each bar of the activity motif, which
    # is how the chart shows hatch; `span-delay` is the same motif collapsed to one
    # bar whose second half has slipped. `procurement` and `span-delay` are
    # deterministic compositions of Lucide's 24 x 24 bar metrics, not traced
    # screenshots. `circle-dashed` is Lucide's own draft ring, adopted on the owner's
    # live review after `pen` was rejected: at 16 px the pen's nib read as a padlock
    # beside the two real padlock glyphs, which is the worst possible confusion on a
    # tab where "locked" is a state.
    'procurement'       = @(
        'M6 4h12'
        'M6 7h12'
        'M4 11h10'
        'M4 14h10'
        'M12 18h8'
        'M12 21h8'
    )
    'span-delay'        = @(
        'M3 10h9'
        'M12 15h9'
    )
    'circle-dashed'     = @(
        'M10.1 2.182a10 10 0 0 1 3.8 0'
        'M13.9 21.818a10 10 0 0 1-3.8 0'
        'M17.609 3.721a10 10 0 0 1 2.69 2.7'
        'M2.182 13.9a10 10 0 0 1 0-3.8'
        'M20.279 17.609a10 10 0 0 1-2.7 2.69'
        'M21.818 10.1a10 10 0 0 1 0 3.8'
        'M3.721 6.391a10 10 0 0 1 2.7-2.69'
        'M6.391 20.279a10 10 0 0 1-2.69-2.7'
    )
    # The two critical rows keep their family shape and gain ONE shared gesture: a
    # wide diagonal struck through it, which is how the chart shows a critical
    # interval overlaying a parent bar. Two slugs rather than one because a bar
    # glyph on a milestone row would name the wrong shape, and the diagonal alone
    # would say "something is critical" without saying what.
    'critical-span'     = @(
        'M6 5h10'
        'M4 12h9'
        'M9 19h9'
        'M6 20 18 4'
    )
    'critical-milestone' = @(
        'M12 13v8'
        'M12 3v3'
        'M4 6a1 1 0 0 0-1 1v5a1 1 0 0 0 1 1h13a2 2 0 0 0 1.152-.365l3.424-2.317a1 1 0 0 0 0-1.635l-3.424-2.318A2 2 0 0 0 17 6z'
        'M6 20 18 4'
    )
    'delineator'        = @(
        'M12 3v18'
        'm16 16 4-4-4-4'
        'm8 8-4 4 4 4'
    )
    'spacer'            = @(
        'M12 2v20'
        'm8 18 4 4 4-4'
        'm8 6 4-4 4 4'
    )
    'splitter'          = @(
        'M14 10h2'
        'M15 22v-8'
        'M15 2v4'
        'M2 10h2'
        'M20 10h2'
        'M3 19h18'
        'M3 22v-6a2 2 135 0 1 2-2h14a2 2 45 0 1 2 2v6'
        'M3 2v2a2 2 45 0 0 2 2h14a2 2 135 0 0 2-2V2'
        'M8 10h2'
        'M9 22v-8'
        'M9 2v4'
    )
    'remove'            = @(
        'M10 11v6'
        'M14 11v6'
        'M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6'
        'M3 6h18'
        'M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2'
    )
    'promote'           = @(
        'M21 5H11'
        'M21 12H11'
        'M21 19H11'
        'm7 8-4 4 4 4'
    )
    'demote'            = @(
        'M21 5H11'
        'M21 12H11'
        'M21 19H11'
        'm3 8 4 4-4 4'
    )
    'move-up'           = @(
        'm5 12 7-7 7 7'
        'M12 19V5'
    )
    'move-down'         = @(
        'M12 5v14'
        'm19 12-7 7-7-7'
    )
    'expand-all'        = @(
        'm7 15 5 5 5-5'
        'm7 9 5-5 5 5'
    )
    'collapse-all'      = @(
        'm7 20 5-5 5 5'
        'm7 4 5 5 5-5'
    )
    'show-columns'      = @(
        'rect:3,3,18,18,2'
        'M9 3v18'
        'M15 3v18'
    )
    'column-visible'    = @(
        'M20 6 9 17 l-5 -5'
    )
    'column-hidden'     = @(
        'M5 12 h14'
    )
    'theme-settings'    = @(
        'M12 22a1 1 0 0 1 0-20 10 9 0 0 1 10 9 5 5 0 0 1-5 5h-2.25a1.75 1.75 0 0 0-1.4 2.8l.3.4a1.75 1.75 0 0 1-1.4 2.8z'
        'circle:13.5,6.5,0.5,fill'
        'circle:17.5,10.5,0.5,fill'
        'circle:6.5,12.5,0.5,fill'
        'circle:8.5,7.5,0.5,fill'
    )
    'chart-title'       = @(
        'M12 4v16'
        'M4 7V5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v2'
        'M9 20h6'
    )
    # R5.1a: the plot-date pickers. Lucide `calendar`, shared by both rows:
    # the label column and the supertip name the end, so two glyphs would add
    # an asset pair and a second decode for no gain.
    'calendar'          = @(
        'M8 2v4'
        'M16 2v4'
        'rect:3,4,18,18,2'
        'M3 10h18'
    )
    'refresh-chart'     = @(
        'M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8'
        'M21 3v5h-5'
        'M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16'
        'M8 16H3v5'
    )
    'warnings'          = @(
        'm21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3'
        'M12 9v4'
        'M12 17h.01'
    )
    'cell-lock-closed'  = @(
        'rect:3,11,18,11,2'
        'M7 11V7a5 5 0 0 1 10 0v4'
    )
    'cell-lock-open'    = @(
        'rect:3,11,18,11,2'
        'M7 11V7a5 5 0 0 1 9.9-1'
    )
    'chart-lock-closed' = @(
        'circle:12,16,1'
        'rect:3,10,18,12,2'
        'M7 10V7a5 5 0 0 1 10 0v3'
    )
    'chart-lock-open'   = @(
        'circle:12,16,1'
        'rect:3,10,18,12,2'
        'M7 10V7a5 5 0 0 1 9.33-2.5'
    )
    'diagnostics'       = @(
        'm12 14 4-4'
        'M3.34 19a10 10 0 1 1 17.32 0'
    )
    'open-log'          = @(
        'M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z'
        'M14 2v5a1 1 0 0 0 1 1h5'
        'M10 9H8'
        'M16 13H8'
        'M16 17H8'
    )
}

# Control id -> glyph slug and pixel size. This is the single source for what the
# ribbon consumes; the generated C# catalogue is projected from the same table so
# the two can never disagree.
$controlGlyphs = [ordered]@{
    btnInitialiseSheet      = @{ Glyph = 'initialise-sheet'; Size = 32 }
    btnRepairConfig         = @{ Glyph = 'repair'; Size = 32 }
    btnValidateSheet        = @{ Glyph = 'validate'; Size = 16 }
    entAsPlannedActivity    = @{ Glyph = 'add-activity'; Size = 16 }
    entAsBuiltActivity      = @{ Glyph = 'add-activity'; Size = 16 }
    entBaselineActivity     = @{ Glyph = 'add-activity'; Size = 16 }
    entCustomActivity       = @{ Glyph = 'circle-dashed'; Size = 16 }
    entDelayEvent           = @{ Glyph = 'span-delay'; Size = 16 }
    entCriticalInterval     = @{ Glyph = 'critical-span'; Size = 16 }
    entAsPlannedProcurement = @{ Glyph = 'procurement'; Size = 16 }
    entAsBuiltProcurement   = @{ Glyph = 'procurement'; Size = 16 }
    entBaselineProcurement  = @{ Glyph = 'procurement'; Size = 16 }
    entAsPlannedMilestone   = @{ Glyph = 'add-milestone'; Size = 16 }
    entAsBuiltMilestone     = @{ Glyph = 'add-milestone'; Size = 16 }
    entBaselineMilestone    = @{ Glyph = 'add-milestone'; Size = 16 }
    entCriticalMilestone    = @{ Glyph = 'critical-milestone'; Size = 16 }
    entCustomMilestone      = @{ Glyph = 'circle-dashed'; Size = 16 }
    entDelineator           = @{ Glyph = 'delineator'; Size = 16 }
    entSpacer               = @{ Glyph = 'spacer'; Size = 16 }
    entSplitter             = @{ Glyph = 'splitter'; Size = 16 }
    btnRemove               = @{ Glyph = 'remove'; Size = 32 }
    btnPromoteToParent      = @{ Glyph = 'promote'; Size = 16 }
    btnDemoteToChild        = @{ Glyph = 'demote'; Size = 16 }
    btnMoveUp               = @{ Glyph = 'move-up'; Size = 16 }
    btnMoveDown             = @{ Glyph = 'move-down'; Size = 16 }
    btnExpandAllGroups      = @{ Glyph = 'expand-all'; Size = 16 }
    btnCollapseAllGroups    = @{ Glyph = 'collapse-all'; Size = 16 }
    btnColumnVisibilityAll  = @{ Glyph = 'show-columns'; Size = 32 }
    btnThemeSettings        = @{ Glyph = 'theme-settings'; Size = 16 }
    btnChartTitle           = @{ Glyph = 'chart-title'; Size = 16 }
    btnPlotStartCalendar    = @{ Glyph = 'calendar'; Size = 16 }
    btnPlotFinishCalendar   = @{ Glyph = 'calendar'; Size = 16 }
    btnRefreshSheet         = @{ Glyph = 'refresh-chart'; Size = 32 }
    btnWarnings             = @{ Glyph = 'warnings'; Size = 16 }
    btnDiagnostics          = @{ Glyph = 'diagnostics'; Size = 16 }
    btnOpenLog              = @{ Glyph = 'open-log'; Size = 16 }
}

# Pictures that follow state rather than control identity, as `slug|size`. The three
# add primaries are here because their label follows the last menu pick (R5.18g): a
# button that says "Spacer" while drawing a delineator is the same defect the label
# fix removed, one channel over.
$stateGlyphs = @(
    'column-visible|16'
    'column-hidden|16'
    'cell-lock-closed|16'
    'cell-lock-open|16'
    'chart-lock-closed|16'
    'chart-lock-open|16'
    'add-activity|32'
    'add-milestone|32'
    'delineator|32'
    'spacer|32'
    'splitter|32'
    'procurement|32'
    'span-delay|32'
    'circle-dashed|32'
    'critical-span|32'
    'critical-milestone|32'
)

# Positive guard: a table entry naming geometry that does not exist is a hard
# failure, not a blank glyph at run time.
$missing = @()
foreach ($entry in $controlGlyphs.Values) {
    if (-not $glyphs.Contains($entry.Glyph)) {
        $missing += $entry.Glyph
    }
}
foreach ($pair in $stateGlyphs) {
    $slug = $pair.Split('|')[0]
    if (-not $glyphs.Contains($slug)) {
        $missing += $slug
    }
}
if ($missing.Count -gt 0) {
    throw "Glyph table names geometry that is not defined: $($missing -join ', ')"
}

function Get-GlyphGeometry {
    [CmdletBinding()]
    param([string]$Descriptor)

    if ($Descriptor.StartsWith('rect:', [StringComparison]::Ordinal)) {
        $parts = $Descriptor.Substring(5).Split(',')
        $bounds = [System.Windows.Rect]::new(
            [double]$parts[0], [double]$parts[1], [double]$parts[2], [double]$parts[3])
        $rect = [System.Windows.Media.RectangleGeometry]::new(
            $bounds, [double]$parts[4], [double]$parts[4])
        return @{ Geometry = $rect; Filled = $false }
    }

    if ($Descriptor.StartsWith('circle:', [StringComparison]::Ordinal)) {
        $parts = $Descriptor.Substring(7).Split(',')
        $circle = [System.Windows.Media.EllipseGeometry]::new(
            [System.Windows.Point]::new([double]$parts[0], [double]$parts[1]),
            [double]$parts[2], [double]$parts[2])
        $filled = $parts.Count -gt 3 -and $parts[3] -eq 'fill'
        return @{ Geometry = $circle; Filled = $filled }
    }

    return @{ Geometry = [System.Windows.Media.Geometry]::Parse($Descriptor); Filled = $false }
}

# Parse every descriptor once, so a malformed path fails here rather than at
# render time inside a loop.
foreach ($slug in $glyphs.Keys) {
    foreach ($descriptor in $glyphs[$slug]) {
        $null = Get-GlyphGeometry -Descriptor $descriptor
    }
}

if (-not (Test-Path -LiteralPath $ResourceDirectory)) {
    $null = New-Item -ItemType Directory -Path $ResourceDirectory -Force
}

$written = 0

function Save-Glyph {
    [CmdletBinding()]
    param(
        [string]$Slug,
        [int]$Size,
        [string]$Palette,
        [System.Windows.Media.Color]$Stroke
    )

    $visual = [System.Windows.Media.DrawingVisual]::new()
    $context = $visual.RenderOpen()
    try {
        $brush = [System.Windows.Media.SolidColorBrush]::new($Stroke)
        $pen = [System.Windows.Media.Pen]::new($brush, $strokeWidth)
        $pen.StartLineCap = [System.Windows.Media.PenLineCap]::Round
        $pen.EndLineCap = [System.Windows.Media.PenLineCap]::Round
        $pen.LineJoin = [System.Windows.Media.PenLineJoin]::Round

        $scale = $Size / $viewBox
        $context.PushTransform([System.Windows.Media.ScaleTransform]::new($scale, $scale))
        foreach ($descriptor in $script:glyphs[$Slug]) {
            $item = Get-GlyphGeometry -Descriptor $descriptor
            if ($item.Filled) {
                $context.DrawGeometry($brush, $null, $item.Geometry)
            }
            else {
                $context.DrawGeometry($null, $pen, $item.Geometry)
            }
        }
        $context.Pop()
    }
    finally {
        $context.Close()
    }

    $bitmap = [System.Windows.Media.Imaging.RenderTargetBitmap]::new(
        $Size, $Size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)

    $encoder = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($bitmap))

    $path = Join-Path $ResourceDirectory ('{0}-{1}-{2}.png' -f $Slug, $Palette, $Size)
    $stream = [System.IO.File]::Create($path)
    try {
        $encoder.Save($stream)
    }
    finally {
        $stream.Dispose()
    }

    return (Get-Item -LiteralPath $path)
}

$pairs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
foreach ($entry in $controlGlyphs.Values) {
    $null = $pairs.Add(('{0}|{1}' -f $entry.Glyph, $entry.Size))
}
foreach ($pair in $stateGlyphs) {
    $null = $pairs.Add($pair)
}

foreach ($pair in $pairs) {
    $parts = $pair.Split('|')
    $slug = $parts[0]
    $size = [int]$parts[1]
    foreach ($palette in $palettes.Keys) {
        $file = Save-Glyph -Slug $slug -Size $size -Palette $palette -Stroke $palettes[$palette]
        Write-Host ('wrote {0} ({1} bytes)' -f $file.Name, $file.Length)
        $written++
    }
}

# The C# catalogue is emitted from the SAME tables that drew the PNGs, so a glyph
# cannot exist in one and not the other. scripts/check-ribbon-glyphs.ps1 re-derives
# this file from the ribbon and fails if the two disagree.
$cataloguePath = Join-Path (Split-Path -Parent $ResourceDirectory) 'RibbonGlyphCatalog.g.cs'

$stateConstants = @(
    '    /// <summary>The tick shown on a column row whose column is visible.</summary>'
    '    internal const string ColumnVisible = "column-visible";'
    ''
    '    /// <summary>The dash shown on a column row whose column is hidden.</summary>'
    '    internal const string ColumnHidden = "column-hidden";'
    ''
    '    /// <summary>The closed padlock shown while the cells are locked.</summary>'
    '    internal const string CellLockClosed = "cell-lock-closed";'
    ''
    '    /// <summary>The open padlock shown while the cells are unlocked.</summary>'
    '    internal const string CellLockOpen = "cell-lock-open";'
    ''
    '    /// <summary>The closed padlock shown while the chart shapes are locked.</summary>'
    '    internal const string ChartLockClosed = "chart-lock-closed";'
    ''
    '    /// <summary>The open padlock shown while the chart shapes are unlocked.</summary>'
    '    internal const string ChartLockOpen = "chart-lock-open";'
)

$builder = [System.Collections.Generic.List[string]]::new()
$builder.Add('// <auto-generated>')
$builder.Add('//     Generated by scripts/gen-ribbon-glyphs.ps1 from the Lucide geometry table.')
$builder.Add('//     Do not edit by hand: edit the script and rerun it.')
$builder.Add('// </auto-generated>')
$builder.Add('#nullable enable')
$builder.Add('')
$builder.Add('using System;')
$builder.Add('using System.Collections.Generic;')
$builder.Add('')
$builder.Add('namespace GanttCreator.AddIn;')
$builder.Add('')
$builder.Add('/// <summary>')
$builder.Add('/// The control-to-glyph table, projected from the same source that draws the PNGs.')
$builder.Add('/// </summary>')
$builder.Add('/// <remarks>')
$builder.Add('/// <para>')
$builder.Add('/// <b>One table, two projections.</b> The generator emits the PNG assets and this file')
$builder.Add('/// from the same declarations, so a control cannot carry a glyph the assets do not')
$builder.Add('/// contain. <c>scripts/check-ribbon-glyphs.ps1</c> re-derives the mapping from the')
$builder.Add('/// shipped ribbon and fails when the two disagree.')
$builder.Add('/// </para>')
$builder.Add('/// <para>')
$builder.Add('/// <b>Sizes are declared, not inferred.</b> A menu row is a normal-height row, so it')
$builder.Add('/// asks for 16 px; a large control asks for 32 px. The callback cannot know which it')
$builder.Add('/// is answering, so the table says so once.')
$builder.Add('/// </para>')
$builder.Add('/// </remarks>')
$builder.Add('internal static class RibbonGlyphCatalog')
$builder.Add('{')
foreach ($line in $stateConstants) {
    $builder.Add($line)
}
$builder.Add('')
$builder.Add('    /// <summary>The glyph of one control: the asset slug and the pixel size to serve.</summary>')
$builder.Add('    internal readonly struct Entry')
$builder.Add('    {')
$builder.Add('        /// <summary>Creates an entry.</summary>')
$builder.Add('        /// <param name="glyph">The asset slug, without palette or size.</param>')
$builder.Add('        /// <param name="size">The pixel size to serve: 16 or 32.</param>')
$builder.Add('        internal Entry(string glyph, int size)')
$builder.Add('        {')
$builder.Add('            Glyph = glyph;')
$builder.Add('            Size = size;')
$builder.Add('        }')
$builder.Add('')
$builder.Add('        /// <summary>Gets the asset slug.</summary>')
$builder.Add('        internal string Glyph { get; }')
$builder.Add('')
$builder.Add('        /// <summary>Gets the pixel size to serve.</summary>')
$builder.Add('        internal int Size { get; }')
$builder.Add('    }')
$builder.Add('')
$builder.Add('    private static readonly Dictionary<string, Entry> _controls = new(StringComparer.Ordinal)')
$builder.Add('    {')
foreach ($id in $controlGlyphs.Keys) {
    $entry = $controlGlyphs[$id]
    $builder.Add(('        ["{0}"] = new Entry("{1}", {2}),' -f $id, $entry.Glyph, $entry.Size))
}
$builder.Add('    };')
$builder.Add('')
$builder.Add('    /// <summary>Gets every mapped control, keyed by its RibbonX id.</summary>')
$builder.Add('    internal static IReadOnlyDictionary<string, Entry> Controls => _controls;')
$builder.Add('')
$builder.Add('    /// <summary>Resolves one control id to its glyph, never throwing.</summary>')
$builder.Add('    /// <param name="controlId">The RibbonX control id, or null.</param>')
$builder.Add('    /// <param name="entry">The resolved entry when the id is mapped.</param>')
$builder.Add('    /// <returns>True when the id is mapped; false otherwise.</returns>')
$builder.Add('    internal static bool TryGet(string? controlId, out Entry entry)')
$builder.Add('    {')
$builder.Add('        if (controlId is not null && _controls.TryGetValue(controlId, out Entry found))')
$builder.Add('        {')
$builder.Add('            entry = found;')
$builder.Add('            return true;')
$builder.Add('        }')
$builder.Add('')
$builder.Add('        entry = default;')
$builder.Add('        return false;')
$builder.Add('    }')
$builder.Add('}')

Set-Content -LiteralPath $cataloguePath -Value $builder -Encoding utf8NoBOM
Write-Host ('wrote {0}' -f (Split-Path -Leaf $cataloguePath))

Write-Host ('{0} glyph image(s) written to {1}' -f $written, $ResourceDirectory)
