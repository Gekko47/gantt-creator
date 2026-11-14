#requires -Version 7
<#
.SYNOPSIS
    Regenerates the two column-visibility item images from the Lucide glyph geometry.

.DESCRIPTION
    The ribbon's column dropdown states each column's visibility with an item image
    (R5.18h): a tick when the column is shown, a dash when it is hidden. RibbonX wants an
    IPictureDisp, and the add-in carries no SVG renderer, so the glyphs are rasterised
    HERE - once, under review - and committed as PNG assets. The runtime only decodes two
    embedded PNGs; nothing renders a path at run time.

    The geometry is Lucide's, verbatim, on Lucide's 24 x 24 view box:
      check  M20 6 9 17 l-5 -5   -> (20,6) (9,17) (4,12)
      minus  M5 12 h14           -> (5,12) (19,12)
    Stroke width 2, round caps and joins, as Lucide specifies. Only the size changes.

    Sizes are 16 px and 32 px, the two Office asks for (small and large DPI).

.PARAMETER OutputDirectory
    Where the PNGs are written. Defaults to the add-in's Resources folder.

.EXAMPLE
    pwsh -NoProfile -File scripts/gen-column-visibility-icons.ps1

.NOTES
    Deterministic: the same script produces byte-identical files, so a diff means the
    geometry or the size changed. Run it, eyeball the two PNGs, commit both.
#>
[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\src\GanttCreator.AddIn\Resources')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing

$viewBox = 24.0
$strokeWidth = 2.0
$sizes = @(16, 32)

# Lucide geometry, in view-box units.
$glyphs = [ordered]@{
    'column-visible' = @(
        [System.Drawing.PointF]::new(20.0, 6.0),
        [System.Drawing.PointF]::new(9.0, 17.0),
        [System.Drawing.PointF]::new(4.0, 12.0)
    )
    'column-hidden' = @(
        [System.Drawing.PointF]::new(5.0, 12.0),
        [System.Drawing.PointF]::new(19.0, 12.0)
    )
}

if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    $null = New-Item -ItemType Directory -Path $OutputDirectory -Force
}

foreach ($size in $sizes) {
    foreach ($name in $glyphs.Keys) {
        $path = Join-Path $OutputDirectory ("{0}-{1}.png" -f $name, $size)
        $bitmap = [System.Drawing.Bitmap]::new($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        try {
            $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
            try {
                $graphics.Clear([System.Drawing.Color]::Transparent)
                $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
                $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
                $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

                # One transform for everything: the pen width is in view-box units, so
                # the scale multiplies the stroke as well as the path. Writing the width
                # pre-divided would put the stroke back at its unscaled size.
                $scale = [float]($size / $viewBox)
                $graphics.ScaleTransform($scale, $scale)

                $pen = [System.Drawing.Pen]::new(
                    [System.Drawing.Color]::FromArgb(255, 0, 0, 0),
                    [float]$strokeWidth)
                try {
                    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
                    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
                    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
                    $graphics.DrawLines($pen, [System.Drawing.PointF[]]$glyphs[$name])
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

        $written = Get-Item -LiteralPath $path
        Write-Host ("wrote {0} ({1} bytes)" -f $written.Name, $written.Length)
    }
}
