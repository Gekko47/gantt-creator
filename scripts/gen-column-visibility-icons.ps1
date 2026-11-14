#requires -Version 7
<#
.SYNOPSIS
    Regenerates the ribbon's committed glyph images from the Lucide geometry.

.DESCRIPTION
    Two families use this script, and both serve a picture from a callback rather than a
    gallery id:

      * the column dropdown's item images (R5.18h): a tick when the column is shown, a dash
        when it is hidden;
      * the two lock toggles' own glyphs (R5.18f, owner ruling D7 2026-11-14): a closed
        padlock while the lock is applied, an open one while it is released. That pairing is
        what the original `imageMso` design could not have, because no open-padlock id exists
        in the Office gallery.

    RibbonX wants an IPictureDisp and the add-in carries no SVG renderer, so the glyphs are
    rasterised HERE - once, under review - and committed as PNG assets. The runtime only
    decodes embedded PNGs; nothing renders a path or an arc at run time.

    The geometry is Lucide's, on Lucide's 24 x 24 view box, stroke width 2, round caps and
    joins, as Lucide specifies. Only the size changes:

      check       M20 6 9 17 l-5 -5                  -> polyline (20,6) (9,17) (4,12)
      minus       M5 12 h14                          -> polyline (5,12) (19,12)
      lock        rect x=3 y=11 w=18 h=11 rx=2       -> rounded body
                  M7 11V7a5 5 0 0 1 10 0v4           -> shackle: half circle to the right side
      lock-open   rect x=3 y=11 w=18 h=11 rx=2       -> rounded body
                  M7 11V7a5 5 0 0 1 9.9-1            -> shackle: stops short of the right side

    Sizes are 16 px and 32 px, the two Office asks for (small and large DPI).

.PARAMETER OutputDirectory
    Where the PNGs are written. Defaults to the add-in's Resources folder.

.EXAMPLE
    pwsh -NoProfile -File scripts/gen-column-visibility-icons.ps1

.NOTES
    The file name still says "column-visibility" because that is the family it was written
    for and the name is referenced from the code and the docs; the lock pair was added to the
    same script rather than to a second, near-identical one.

    Deterministic: the same script produces byte-identical files, so a diff means the
    geometry or the size changed. Run it, eyeball the PNGs, commit them.
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

# A polyline, in view-box units. `Get-` because the function returns a drawing, it does not
# change state - PSScriptAnalyzer's PSUseShouldProcessForStateChangingFunctions rejects `New-`
# for exactly that reason.
function Get-PolylineDrawing {
    param([System.Drawing.PointF[]]$Points)
    return {
        param($Graphics, $Pen)
        $Graphics.DrawLines($Pen, $Points)
    }.GetNewClosure()
}

# Lucide's rounded lock body: an 18 x 11 rect with 2-unit corners, drawn as one closed path.
function Get-LockBodyDrawing {
    return {
        param($Graphics, $Pen)
        $diameter = 4.0
        $left = 3.0
        $top = 11.0
        $right = 21.0
        $bottom = 22.0
        $body = [System.Drawing.Drawing2D.GraphicsPath]::new()
        try {
            # Angles run clockwise from 3 o'clock in GDI+, so each corner starts at its own
            # start angle and sweeps 90 degrees.
            $body.AddArc($left, $top, $diameter, $diameter, 180, 90)
            $body.AddArc(($right - $diameter), $top, $diameter, $diameter, 270, 90)
            $body.AddArc(($right - $diameter), ($bottom - $diameter), $diameter, $diameter, 0, 90)
            $body.AddArc($left, ($bottom - $diameter), $diameter, $diameter, 90, 90)
            $body.CloseFigure()
            $Graphics.DrawPath($Pen, $body)
        }
        finally {
            $body.Dispose()
        }
    }.GetNewClosure()
}

# Lucide's shackle: a half circle of radius 5 centred on the body's top edge. The open lock
# stops short of the right-hand side, which is the whole visual difference between the pair.
function Get-ShackleDrawing {
    param([int]$Sweep)
    return {
        param($Graphics, $Pen)
        $shackle = [System.Drawing.Drawing2D.GraphicsPath]::new()
        try {
            $shackle.AddArc(7.0, 2.0, 10.0, 10.0, 180, $Sweep)
            $Graphics.DrawPath($Pen, $shackle)
        }
        finally {
            $shackle.Dispose()
        }
    }.GetNewClosure()
}

$glyphs = [ordered]@{
    'column-visible' = Get-PolylineDrawing -Points @(
        [System.Drawing.PointF]::new(20.0, 6.0),
        [System.Drawing.PointF]::new(9.0, 17.0),
        [System.Drawing.PointF]::new(4.0, 12.0)
    )
    'column-hidden'  = Get-PolylineDrawing -Points @(
        [System.Drawing.PointF]::new(5.0, 12.0),
        [System.Drawing.PointF]::new(19.0, 12.0)
    )
    'lock-locked'    = @((Get-LockBodyDrawing), (Get-ShackleDrawing -Sweep 180))
    'lock-unlocked'  = @((Get-LockBodyDrawing), (Get-ShackleDrawing -Sweep 150))
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

                    # A glyph is either one drawing action or a list of them.
                    $drawings = @($glyphs[$name])
                    if ($drawings.Count -eq 1 -and $drawings[0] -is [scriptblock]) {
                        & $drawings[0] $graphics $pen
                    }
                    else {
                        foreach ($drawing in $drawings) {
                            & $drawing $graphics $pen
                        }
                    }
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
