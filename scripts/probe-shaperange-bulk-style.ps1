#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: does a grouped Excel ShapeRange style write cost less than the
    same writes issued shape by shape, and does it land the same values?

.DESCRIPTION
    The live refresh re-applies fill, stroke, typography, and text formatting one shape
    at a time. A full update is roughly 20-40 COM calls per shape, and the measured write
    phase is 86 % of a full manual rebuild. The candidate reduction is to group the
    shapes by identical appearance and apply the secondary format objects (Fill, Line,
    TextFrame2) to one ShapeRange per group. Microsoft documents that setting a property
    of those secondary objects on a ShapeRange applies it to every shape in the range.

    This probe answers three questions on a disposable workbook:

      1. How long does Shapes.Range(names) take to construct for 15-60 names?
      2. Per-shape style loop vs one ShapeRange call per appearance group: measured cost.
      3. Does every shape read back the grouped values, including text frame margins,
         vertical anchor, font size/bold/name, paragraph alignment, and font fill colour?

    It also tries a TextFrame2 write on a mixed range (a rectangle plus a text box),
    because the live writer must group by kind if the host refuses that combination.

    The probe creates and discards its own workbook in a private, invisible Excel
    instance. It never touches an open user workbook.

.EXAMPLE
    pwsh ./scripts/probe-shaperange-bulk-style.ps1

.NOTES
    Disposable spike, in the style of probe-plot-origin.ps1. Results are printed as they
    are measured because COM proxy values do not survive a PowerShell function boundary
    reliably. The script changes no production code.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$rectGroups = @(
    @{ Name = 'rect-a'; Count = 20; Fill = 0x00EFC6; Line = 0x008000; Weight = 0.75; Dash = 1; Transparency = 0.0 }
    @{ Name = 'rect-b'; Count = 20; Fill = 0x00C7FF; Line = 0xC00000; Weight = 1.0; Dash = 4; Transparency = 0.1 }
    @{ Name = 'rect-c'; Count = 20; Fill = 0x00EBDD; Line = 0x1F4E79; Weight = 1.25; Dash = 7; Transparency = 0.0 }
)

$textGroups = @(
    @{ Name = 'text-a'; Count = 15; Fill = 0x00FFFFFF; Line = 0x404040; Weight = 0.5; Dash = 1; Transparency = 0.0; FontSize = 8.0; Bold = -1; Alignment = 2; TextColour = 0x404040; Margin = 2.0 }
    @{ Name = 'text-b'; Count = 15; Fill = 0x00F2F2F2; Line = 0x000000; Weight = 0.75; Dash = 1; Transparency = 0.0; FontSize = 10.0; Bold = 0; Alignment = 1; TextColour = 0x000000; Margin = 0.0 }
)

$iterations = 5
$script:warnings = [System.Collections.Generic.List[string]]::new()
$script:totalMismatches = 0

function Test-CloseEnough {
    param(
        [Parameter(Mandatory)] [double] $Actual,
        [Parameter(Mandatory)] [double] $Expected
    )

    return [Math]::Abs($Actual - $Expected) -le 0.01
}

function Write-ProbeLine {
    param([Parameter(Mandatory)] [AllowEmptyString()] [string] $Text)

    Write-Host $Text
}

function Get-Median {
    param([Parameter(Mandatory)] [double[]] $Values)

    $sorted = @($Values | Sort-Object)
    $count = $sorted.Count
    if (($count % 2) -eq 1) {
        return [double]$sorted[[int]($count / 2)]
    }

    return ([double]$sorted[($count / 2) - 1] + [double]$sorted[$count / 2]) / 2.0
}

function Get-ShapeRange {
    param(
        [Parameter(Mandatory)] $Shapes,
        [Parameter(Mandatory)] [string[]] $Names
    )

    # The unary comma prevents PowerShell from enumerating the returned ShapeRange into
    # an array of Shape objects. An enumerated range has no Fill/Line/TextFrame2, and the
    # first property write then fails against an array rather than against the host.
    try {
        return , $Shapes.Range([object[]]$Names)
    }
    catch {
        Write-ProbeLine ('  Shapes.Range(object[]) failed: ' + $_.Exception.Message)
        return , $Shapes.Range($Names)
    }
}

function Write-RectStyle {
    param(
        [Parameter(Mandatory)] $Target,
        [Parameter(Mandatory)] [hashtable] $Group
    )

    # Resolve the secondary objects INSIDE the function, from the shape or range the
    # caller passed. Passing a COM proxy as a parameter loses its late-bound type
    # information in PowerShell, and the first property write then reports "RGB cannot
    # be found" against an otherwise valid FillFormat.
    $Fill = $Target.Fill
    $Line = $Target.Line
    $Fill.ForeColor.RGB = $Group.Fill
    $Fill.Transparency = $Group.Transparency
    $Fill.Visible = -1
    $Line.ForeColor.RGB = $Group.Line
    $Line.Weight = $Group.Weight
    $Line.DashStyle = $Group.Dash
    $Line.Transparency = $Group.Transparency
    $Line.Visible = -1
}

function Write-TextStyle {
    param(
        [Parameter(Mandatory)] $Target,
        [Parameter(Mandatory)] [hashtable] $Group
    )

    $TextFrame = $Target.TextFrame2
    $TextFrame.WordWrap = 0
    $TextFrame.AutoSize = 0
    $TextFrame.VerticalAnchor = 3
    $TextFrame.MarginLeft = $Group.Margin
    $TextFrame.MarginRight = 0.0
    $TextFrame.MarginTop = 0.0
    $TextFrame.MarginBottom = 0.0
    $textRange = $TextFrame.TextRange
    $textRange.ParagraphFormat.Alignment = $Group.Alignment
    $font = $textRange.Font
    $font.Name = 'Aptos'
    $font.Size = $Group.FontSize
    $font.Bold = $Group.Bold
    $font.Fill.Solid()
    $font.Fill.ForeColor.RGB = $Group.TextColour
    $font.Fill.Visible = -1
}

function Measure-PerShape {
    param(
        [Parameter(Mandatory)] [string[]] $Names,
        [Parameter(Mandatory)] [hashtable] $Group,
        [Parameter(Mandatory)] [bool] $IsText,
        [Parameter(Mandatory)] $Sheet
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($name in $Names) {
        $shape = $Sheet.Shapes.Item($name)
        Write-RectStyle -Target $shape -Group $Group
        if ($IsText) {
            Write-TextStyle -Target $shape -Group $Group
        }
    }

    $stopwatch.Stop()
    return $stopwatch.Elapsed.TotalMilliseconds
}

function Measure-Grouped {
    param(
        [Parameter(Mandatory)] [string[]] $Names,
        [Parameter(Mandatory)] [hashtable] $Group,
        [Parameter(Mandatory)] [bool] $IsText,
        [Parameter(Mandatory)] $Sheet
    )

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $range = Get-ShapeRange -Shapes $Sheet.Shapes -Names $Names
    Write-RectStyle -Target $range -Group $Group
    if ($IsText) {
        Write-TextStyle -Target $range -Group $Group
    }

    $stopwatch.Stop()
    return $stopwatch.Elapsed.TotalMilliseconds
}

function Write-Baseline {
    param(
        [Parameter(Mandatory)] [string[]] $Names,
        [Parameter(Mandatory)] [hashtable] $Group,
        [Parameter(Mandatory)] [bool] $IsText,
        [Parameter(Mandatory)] $Sheet
    )

    $baseline = @{
        Fill         = 0x00FFFFFF
        Line         = 0x00808080
        Weight       = 0.5
        Dash         = 1
        Transparency = 0.0
        FontSize     = 8.0
        Bold         = 0
        Alignment    = 1
        TextColour   = 0x000000
        Margin       = 0.0
    }
    Measure-PerShape -Names $Names -Group $baseline -IsText $IsText -Sheet $Sheet | Out-Null
}

function Test-GroupEquivalence {
    param(
        [Parameter(Mandatory)] [string[]] $Names,
        [Parameter(Mandatory)] [hashtable] $Group,
        [Parameter(Mandatory)] [bool] $IsText,
        [Parameter(Mandatory)] $Sheet
    )

    $mismatches = 0
    foreach ($name in $Names) {
        $shape = $Sheet.Shapes.Item($name)
        if ([int]$shape.Fill.ForeColor.RGB -ne [int]$Group.Fill) { $mismatches++ }
        if ([int]$shape.Line.ForeColor.RGB -ne [int]$Group.Line) { $mismatches++ }
        if (-not (Test-CloseEnough -Actual ([double]$shape.Line.Weight) -Expected ([double]$Group.Weight))) { $mismatches++ }
        if ([int]$shape.Line.DashStyle -ne [int]$Group.Dash) { $mismatches++ }
        if (-not (Test-CloseEnough -Actual ([double]$shape.Fill.Transparency) -Expected ([double]$Group.Transparency))) { $mismatches++ }
        if ($IsText) {
            $frame = $shape.TextFrame2
            if (-not (Test-CloseEnough -Actual ([double]$frame.MarginLeft) -Expected ([double]$Group.Margin))) { $mismatches++ }
            if ([int]$frame.VerticalAnchor -ne 3) { $mismatches++ }
            $textRange = $frame.TextRange
            if (-not (Test-CloseEnough -Actual ([double]$textRange.Font.Size) -Expected ([double]$Group.FontSize))) { $mismatches++ }
            if ([int]$textRange.Font.Bold -ne [int]$Group.Bold) { $mismatches++ }
            if ([string]$textRange.Font.Name -ne 'Aptos') { $mismatches++ }
            if ([int]$textRange.ParagraphFormat.Alignment -ne [int]$Group.Alignment) { $mismatches++ }
            if ([int]$textRange.Font.Fill.ForeColor.RGB -ne [int]$Group.TextColour) { $mismatches++ }
        }
    }

    return $mismatches
}

function Add-ProbeShapes {
    param(
        [Parameter(Mandatory)] $Sheet,
        [Parameter(Mandatory)] [hashtable] $Group,
        [Parameter(Mandatory)] [bool] $IsText,
        [Parameter(Mandatory)] [int] $StartIndex
    )

    $names = [System.Collections.Generic.List[string]]::new()
    for ($index = 0; $index -lt $Group.Count; $index++) {
        $left = 10.0 + (($StartIndex + $index) % 10) * 50.0
        $top = 10.0 + [Math]::Floor(($StartIndex + $index) / 10) * 20.0
        if ($IsText) {
            $shape = $Sheet.Shapes.AddTextbox(1, $left, $top, 40.0, 12.0)
        }
        else {
            $shape = $Sheet.Shapes.AddShape(1, $left, $top, 40.0, 12.0)
        }

        $shape.Name = $Group.Name + '-' + $index
        $names.Add($shape.Name)
    }

    return $names.ToArray()
}

$excel = $null
$workbook = $null
$sheet = $null
try {
    Write-ProbeLine '=== Probe: grouped ShapeRange style writes ==='
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'ShapeRangeProbe'

    $rectNamesByGroup = @{}
    $textNamesByGroup = @{}
    $index = 0
    foreach ($group in $rectGroups) {
        $rectNamesByGroup[$group.Name] = Add-ProbeShapes -Sheet $sheet -Group $group -IsText $false -StartIndex $index
        $index += $group.Count
    }

    foreach ($group in $textGroups) {
        $textNamesByGroup[$group.Name] = Add-ProbeShapes -Sheet $sheet -Group $group -IsText $true -StartIndex $index
        $index += $group.Count
    }

    Write-ProbeLine ('Created {0} rectangle(s) and {1} text box(es).' -f `
            (($rectNamesByGroup.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum),
        (($textNamesByGroup.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum))
    Write-ProbeLine ''

    $allNames = @($rectNamesByGroup.Values) + @($textNamesByGroup.Values)
    foreach ($groupNames in $allNames) {
        $rangeSizes = @($groupNames).Count
        $timings = [System.Collections.Generic.List[double]]::new()
        for ($round = 0; $round -lt 10; $round++) {
            $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $range = $sheet.Shapes.Range([object[]]$groupNames)
            $null = $range.Count
            $stopwatch.Stop()
            $timings.Add($stopwatch.Elapsed.TotalMilliseconds)
        }

        Write-ProbeLine ('Shapes.Range n={0}: median {1:0.###} ms' -f $rangeSizes, (Get-Median -Values $timings.ToArray()))
    }

    Write-ProbeLine ''
    foreach ($group in $rectGroups) {
        $names = $rectNamesByGroup[$group.Name]
        Write-Baseline -Names $names -Group $group -IsText $false -Sheet $sheet
        $timings = [System.Collections.Generic.List[double]]::new()
        for ($round = 0; $round -lt $iterations; $round++) {
            $timings.Add((Measure-PerShape -Names $names -Group $group -IsText $false -Sheet $sheet))
            Write-Baseline -Names $names -Group $group -IsText $false -Sheet $sheet
        }

        Write-ProbeLine ('Per-shape rect {0} n={1}: median {2:0.###} ms ({3:0.###} ms/shape)' -f `
                $group.Name, $names.Count, (Get-Median -Values $timings.ToArray()),
            ((Get-Median -Values $timings.ToArray()) / $names.Count))
    }

    foreach ($group in $textGroups) {
        $names = $textNamesByGroup[$group.Name]
        Write-Baseline -Names $names -Group $group -IsText $true -Sheet $sheet
        $timings = [System.Collections.Generic.List[double]]::new()
        for ($round = 0; $round -lt $iterations; $round++) {
            $timings.Add((Measure-PerShape -Names $names -Group $group -IsText $true -Sheet $sheet))
            Write-Baseline -Names $names -Group $group -IsText $true -Sheet $sheet
        }

        Write-ProbeLine ('Per-shape text {0} n={1}: median {2:0.###} ms ({3:0.###} ms/shape)' -f `
                $group.Name, $names.Count, (Get-Median -Values $timings.ToArray()),
            ((Get-Median -Values $timings.ToArray()) / $names.Count))
    }

    Write-ProbeLine ''
    foreach ($group in $rectGroups) {
        $names = $rectNamesByGroup[$group.Name]
        Write-Baseline -Names $names -Group $group -IsText $false -Sheet $sheet
        $timings = [System.Collections.Generic.List[double]]::new()
        for ($round = 0; $round -lt $iterations; $round++) {
            $timings.Add((Measure-Grouped -Names $names -Group $group -IsText $false -Sheet $sheet))
            Write-Baseline -Names $names -Group $group -IsText $false -Sheet $sheet
        }

        # The timing loop resets after every round, so apply the group once more before
        # the read-back; without this the probe compares the baseline with the group.
        $null = Measure-Grouped -Names $names -Group $group -IsText $false -Sheet $sheet
        $mismatches = Test-GroupEquivalence -Names $names -Group $group -IsText $false -Sheet $sheet
        $script:totalMismatches += $mismatches
        Write-ProbeLine ('Grouped rect {0} n={1}: median {2:0.###} ms; read-back mismatches {3}' -f `
                $group.Name, $names.Count, (Get-Median -Values $timings.ToArray()), $mismatches)
    }

    foreach ($group in $textGroups) {
        $names = $textNamesByGroup[$group.Name]
        Write-Baseline -Names $names -Group $group -IsText $true -Sheet $sheet
        $timings = [System.Collections.Generic.List[double]]::new()
        for ($round = 0; $round -lt $iterations; $round++) {
            $timings.Add((Measure-Grouped -Names $names -Group $group -IsText $true -Sheet $sheet))
            Write-Baseline -Names $names -Group $group -IsText $true -Sheet $sheet
        }

        $null = Measure-Grouped -Names $names -Group $group -IsText $true -Sheet $sheet
        $mismatches = Test-GroupEquivalence -Names $names -Group $group -IsText $true -Sheet $sheet
        $script:totalMismatches += $mismatches
        Write-ProbeLine ('Grouped text {0} n={1}: median {2:0.###} ms; read-back mismatches {3}' -f `
                $group.Name, $names.Count, (Get-Median -Values $timings.ToArray()), $mismatches)
    }

    Write-ProbeLine ''
    $mixedNames = [string[]]@($rectNamesByGroup['rect-a'][0], $textNamesByGroup['text-a'][0])
    $mixedRange = Get-ShapeRange -Shapes $sheet.Shapes -Names $mixedNames
    try {
        $mixedRange.Fill.ForeColor.RGB = 0x00123456
        Write-ProbeLine 'Mixed rect+text Fill on ShapeRange: accepted.'
    }
    catch {
        Write-ProbeLine ('Mixed rect+text Fill on ShapeRange: refused: ' + $_.Exception.Message)
        $script:warnings.Add('Mixed Fill refused')
    }

    try {
        $mixedRange.TextFrame2.MarginLeft = 3.0
        Write-ProbeLine 'Mixed rect+text TextFrame2 on ShapeRange: accepted.'
    }
    catch {
        Write-ProbeLine ('Mixed rect+text TextFrame2 on ShapeRange: refused: ' + $_.Exception.Message)
        $script:warnings.Add('Mixed TextFrame2 refused')
    }

    Write-ProbeLine ''
    if ($script:warnings.Count -eq 0 -and $script:totalMismatches -eq 0) {
        Write-ProbeLine 'Summary: grouped secondary-format writes are accepted and read back exactly.'
    }
    elseif ($script:warnings.Count -eq 0) {
        Write-ProbeLine ('Summary: grouped writes accepted; ' + $script:totalMismatches + ' read-back mismatch(es) remain.')
    }
    else {
        Write-ProbeLine ('Summary: ' + ($script:warnings -join '; ') + '; ' + $script:totalMismatches + ' read-back mismatch(es).')
    }
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    foreach ($comObject in @($sheet, $workbook, $excel)) {
        if ($null -ne $comObject) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
        }
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
