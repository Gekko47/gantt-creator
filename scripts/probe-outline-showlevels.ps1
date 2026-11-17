#Requires -Version 7.0
<#
.SYNOPSIS
    Disposable probe: how can a Ribbon command collapse and expand ALL the
    outline groups of the Gantt table, and does that survive a Refresh?

.DESCRIPTION
    R5.4's roadmap row asks for "expand/collapse display commands". The native
    Excel `+`/`-` outline controls already work (R4.7D rebuilds the groups from
    `ParentId`; R5.13's gate records that the native controls "reflect it
    without Refresh"), and the owner has ruled that the current visual and
    function are intended. What is missing is a first-class Ribbon surface.

    The Ribbon button must drive the SAME Excel-native state the native control
    drives, so its result is identical to the native control BY CONSTRUCTION.
    That makes the mechanism the whole question, and two candidates were
    measured against a live host:

      CANDIDATE A -- `Range.Outline.ShowLevels(RowLevels, ColumnLevels)`.
        Rejected: `Range.Outline` does not exist on IDispatch at all (the
        PowerShell adapter and a raw `InvokeMember` both report
        DISP_E_UNKNOWNNAME), so the only reachable `Outline` is the
        WORKSHEET's, and `Worksheet.Outline.ShowLevels(1, 1)` is SHEET-WIDE.
        It would collapse the user's own outline groups elsewhere on the sheet
        -- a regression of their content. Rejected on that ground alone.

      CANDIDATE B -- the `Hidden` flag of the child rows, written through
        `EntireRow.Hidden`. This is the SAME state Excel's native `+`/`-`
        controls write (collapsing a group hides its child rows), and
        `ExcelParentChildMutator` already proves this accessor: it records
        that `Range("A3:C3").Hidden = $true` is REFUSED with COMException
        0x800A03EC while the identically-shaped `Range("A3:C3").EntireRow.Hidden
        = $true` is ACCEPTED. It is row-scoped by construction, so it cannot
        touch a group the add-in does not own.

    Candidate B is what this probe measures. Four questions:

      Q1. SCOPE. Does hiding the Gantt table's child rows leave the USER's own
          outline group on the same sheet untouched?

      Q2. REFRESH. Does a Refresh preserve the collapsed state? This is the
          sharp edge. `ExcelOutlineGroupWriter` writes the outline level of
          every planned child row on every Refresh, and its own contract says
          "writing their level re-evaluates the outline and re-expands every
          collapsed group". The guard it applies protects only the LEVEL-1
          rows; a collapsed child is at LEVEL 2 and still hidden, so the group
          pass may rewrite it and pop the group back open. The reset pass
          (level 1 to a row no longer in the plan) is the other half.

      Q3. RESTORE. Does clearing the flag expand every group again?

      Q4. ACCESSOR. Which accessor accepts the write? The one that is
          ACCEPTED here is the one the writer must use; a REFUSED accessor
          must not be relied on.

    Disposable by design: it drives a throwaway workbook through COM and prints
    measurements. It is not a test, is not run by verify.ps1, and its findings
    are recorded as evidence in the R5.4 work item.

.NOTES
    Run with:

        pwsh ./scripts/probe-outline-showlevels.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$bodyHeightPt = 18.0
$headerHeightPt = 16.0

$topLevelOutlineLevel = 1
$childOutlineLevel = 2

function Show-GroupState {
    param($Sheet, [int] $FirstRow, [int] $LastRow, [string] $Label)

    Write-Host ("  {0}" -f $Label)
    foreach ($r in $FirstRow..$LastRow) {
        $row = $Sheet.Rows.Item($r)
        Write-Host (
            "     row {0,-3} OutlineLevel={1,-3} Hidden={2,-6} RowHeight={3}" -f
            $r, $row.OutlineLevel, $row.Hidden, $row.RowHeight)
    }
}

function Invoke-ProbeCall {
    param([string] $What, [scriptblock] $Body)

    try {
        & $Body
        Write-Host ("  {0,-52} ACCEPTED" -f $What)
        return $true
    }
    catch [System.Runtime.InteropServices.COMException] {
        Write-Host ("  {0,-52} REFUSED  COM 0x{1:X8}" -f $What, $_.Exception.HResult)
        return $false
    }
    catch {
        Write-Host ("  {0,-52} ERROR    {1}: {2}" -f $What, $_.Exception.GetType().Name, $_.Exception.Message)
        return $false
    }
}

function Set-RowHidden {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param($Sheet, [int[]] $Rows, [bool] $Hidden)

    foreach ($r in $Rows) {
        if ($PSCmdlet.ShouldProcess("row $r", "set Hidden to $Hidden")) {
            $Sheet.Rows.Item($r).EntireRow.Hidden = $Hidden
        }
    }
}

$excel = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel version: $($excel.Version)"
    Write-Host

    # A sheet shaped like the Gantt sheet:
    #
    #   row 1        header
    #   row 2  P1    parent, outline level 1
    #   row 3  C1    child of P1, outline level 2
    #   row 4  C2    child of P1, outline level 2
    #   row 5  P2    parent, outline level 1
    #   row 6  C3    child of P2, outline level 2
    #
    # plus an UNRELATED group the user owns, on rows 9-10, which the add-in
    # must never touch (Q1).
    $sheet = $excel.Workbooks.Add().Worksheets.Item(1)
    $sheet.Name = 'Outline'
    $table = $sheet.ListObjects.Add(1, $sheet.Range('A1:C6'), $null, 1)
    $table.Name = 'tblGanttData'

    $sheet.Rows.Item(1).RowHeight = $headerHeightPt
    foreach ($r in 2..6) { $sheet.Rows.Item($r).RowHeight = $bodyHeightPt }

    $sheet.Range('A2').Value2 = 'P1'
    $sheet.Range('A3').Value2 = 'C1'
    $sheet.Range('A4').Value2 = 'C2'
    $sheet.Range('A5').Value2 = 'P2'
    $sheet.Range('A6').Value2 = 'C3'

    # The user's own group, outside the Gantt table entirely.
    $sheet.Range('A9').Value2 = 'USER-P'
    $sheet.Range('A10').Value2 = 'USER-C'
    $sheet.Rows.Item(9).RowHeight = $bodyHeightPt
    $sheet.Rows.Item(10).RowHeight = $bodyHeightPt
    $sheet.Rows.Item(10).OutlineLevel = $childOutlineLevel

    # The Gantt groups, written the way `ExcelOutlineGroupWriter` writes them.
    $sheet.Rows.Item(3).OutlineLevel = $childOutlineLevel
    $sheet.Rows.Item(4).OutlineLevel = $childOutlineLevel
    $sheet.Rows.Item(6).OutlineLevel = $childOutlineLevel

    $ganttChildRows = @(3, 4, 6)

    Write-Host '=== BASELINE: the sheet as built (all groups expanded) ==='
    Show-GroupState $sheet 2 6 'Gantt table body (rows 2-6)'
    Show-GroupState $sheet 9 10 "the USER's own group (rows 9-10)"
    Write-Host

    Write-Host '=== CANDIDATE A: Range.Outline.ShowLevels -- is it even reachable? ==='
    Write-Host '  (the PIA declares Range.Outline, but the COM object may not expose it)'
    $null = Invoke-ProbeCall 'table.DataBodyRange.Outline.ShowLevels(1, 1)' {
        $table.DataBodyRange.Outline.ShowLevels(1, 1)
    }
    $null = Invoke-ProbeCall 'worksheet.Outline.ShowLevels(1, 1)  [SHEET-WIDE]' {
        $sheet.Outline.ShowLevels(1, 1)
    }
    Write-Host '  (a worksheet-wide collapse would hide the USER''s own group too,'
    Write-Host '   which is why Candidate A is rejected: see Q1 below)'
    # Undo whatever the sheet-wide call did, so the next case starts clean.
    Set-RowHidden $sheet $ganttChildRows $false
    $sheet.Rows.Item(10).Hidden = $false
    $sheet.Rows.Item(3).OutlineLevel = $childOutlineLevel
    $sheet.Rows.Item(4).OutlineLevel = $childOutlineLevel
    $sheet.Rows.Item(6).OutlineLevel = $childOutlineLevel
    $sheet.Rows.Item(10).OutlineLevel = $childOutlineLevel
    Write-Host

    Write-Host '=== Q4: which accessor accepts the Hidden write? ==='
    $null = Invoke-ProbeCall "Range('A3:C3').Hidden = `$true" { $sheet.Range('A3:C3').Hidden = $true }
    $null = Invoke-ProbeCall "Range('A3:C3').EntireRow.Hidden = `$true" { $sheet.Range('A3:C3').EntireRow.Hidden = $true }
    $null = Invoke-ProbeCall 'Rows.Item(3).EntireRow.Hidden = $false [restore]' { $sheet.Rows.Item(3).EntireRow.Hidden = $false }
    Write-Host '  (the ACCEPTED accessor is the one the writer must use)'
    Write-Host

    Write-Host '=== Q1: does the collapse reach the USER''s own group? ==='
    $null = Invoke-ProbeCall 'hide the Gantt child rows via EntireRow.Hidden' {
        Set-RowHidden $sheet $ganttChildRows $true
    }
    Show-GroupState $sheet 2 6 'Gantt table body after the collapse'
    Show-GroupState $sheet 9 10 "the USER's own group after the collapse"
    $userGroupHidden = [bool]$sheet.Rows.Item(10).Hidden
    Write-Host ("  --> the USER's group was collapsed too? {0}" -f $userGroupHidden)
    Write-Host

    Write-Host '=== Q2: does a REFRESH re-expand a collapsed group? ==='
    # The Refresh GROUP PASS: write the planned child outline level back to
    # every planned child row. The writer''s own contract says "writing their
    # level re-evaluates the outline and re-expands every collapsed group",
    # and its guard protects only the LEVEL-1 rows -- a collapsed child is at
    # LEVEL 2 and still hidden, so this is the write that can pop the group
    # back open.
    foreach ($r in $ganttChildRows) {
        $null = Invoke-ProbeCall ("Rows.Item($r).OutlineLevel = 2  [group pass]") {
            $sheet.Rows.Item($r).OutlineLevel = $childOutlineLevel
        }
    }
    Show-GroupState $sheet 2 6 'Gantt table body after the group pass'
    $groupPassReExpanded = -not [bool]$sheet.Rows.Item(3).Hidden
    Write-Host ("  --> the group pass RE-EXPANDED the collapsed group? {0}" -f $groupPassReExpanded)
    Write-Host

    # The Refresh RESET PASS: write level 1 to a row that is at level 2 and is
    # no longer in the plan. The writer's guard is about exactly this write.
    $null = Invoke-ProbeCall 'Rows.Item(4).OutlineLevel = 1  [reset pass]' {
        $sheet.Rows.Item(4).OutlineLevel = $topLevelOutlineLevel
    }
    Show-GroupState $sheet 2 6 'Gantt table body after the reset pass'
    $resetPassReExpanded = -not [bool]$sheet.Rows.Item(4).Hidden
    Write-Host ("  --> the reset pass RE-EXPANDED the collapsed group? {0}" -f $resetPassReExpanded)
    Write-Host

    # Restore the plan the writer would write, so the expand case starts from
    # the state a Refresh would leave behind.
    $sheet.Rows.Item(4).OutlineLevel = $childOutlineLevel
    Write-Host

    Write-Host '=== Q3: does clearing the flag expand everything again? ==='
    $null = Invoke-ProbeCall 'show the Gantt child rows again' {
        Set-RowHidden $sheet $ganttChildRows $false
    }
    Show-GroupState $sheet 2 6 'Gantt table body after the expand'
    Show-GroupState $sheet 9 10 "the USER's own group after the expand"
    $allExpanded = -not [bool]$sheet.Rows.Item(3).Hidden -and
        -not [bool]$sheet.Rows.Item(4).Hidden -and
        -not [bool]$sheet.Rows.Item(6).Hidden
    Write-Host ("  --> every Gantt group expanded again? {0}" -f $allExpanded)
    Write-Host

    Write-Host '=== Q5: does EXCEL''S OWN collapse produce the state we write? ==='
    # The decisive check for "identical to the native control BY CONSTRUCTION":
    # if Excel's own sheet-wide collapse leaves the child rows Hidden = $true,
    # then writing that flag ourselves produces exactly the state Excel writes,
    # and the Ribbon button is indistinguishable from the native control.
    # A FRESH sheet is used so the sheet-wide call has no other group to touch.
    $nativeSheet = $excel.Workbooks.Add().Worksheets.Item(1)
    $nativeSheet.Name = 'Native'
    $nativeSheet.Rows.Item(2).OutlineLevel = 1
    $nativeSheet.Rows.Item(3).OutlineLevel = 2
    $nativeSheet.Rows.Item(4).OutlineLevel = 2

    Write-Host '  -- Excel''s own collapse: worksheet.Outline.ShowLevels(1, 1) --'
    $null = Invoke-ProbeCall 'nativeSheet.Outline.ShowLevels(1, 1)' {
        $nativeSheet.Outline.ShowLevels(1, 1)
    }
    Show-GroupState $nativeSheet 2 4 'native sheet after Excel''s own collapse'
    $nativeChildHidden = [bool]$nativeSheet.Rows.Item(3).Hidden
    Write-Host ("  --> Excel's own collapse HIDES the child rows? {0}" -f $nativeChildHidden)
    Write-Host '     (a TRUE here means our EntireRow.Hidden write produces exactly'
    Write-Host '      the state Excel writes -- identical by construction)'
    Write-Host

    Write-Host '  -- Excel''s own expand: worksheet.Outline.ShowLevels(2, 1) --'
    $null = Invoke-ProbeCall 'nativeSheet.Outline.ShowLevels(2, 1)' {
        $nativeSheet.Outline.ShowLevels(2, 1)
    }
    Show-GroupState $nativeSheet 2 4 'native sheet after Excel''s own expand'
    $nativeChildShown = -not [bool]$nativeSheet.Rows.Item(3).Hidden
    Write-Host ("  --> Excel's own expand SHOWS the child rows? {0}" -f $nativeChildShown)
    Write-Host

    Write-Host '  -- Q6: does Excel''s expand reveal a row the USER hid by hand? --'
    # A row hidden by hand carries outline level 1 (it belongs to no group).
    # If Excel''s own expand leaves it hidden, then "expand all" must leave
    # hand-hidden rows alone, and the add-in must do the same: it may only
    # reveal the rows ITS relationships hid.
    $nativeSheet.Rows.Item(4).Hidden = $true
    Write-Host '  (row 4 hidden by hand, outline level 2)'
    $null = Invoke-ProbeCall 'nativeSheet.Outline.ShowLevels(2, 1) [expand again]' {
        $nativeSheet.Outline.ShowLevels(2, 1)
    }
    Show-GroupState $nativeSheet 2 4 'native sheet after the expand'
    $handHiddenStayedHidden = [bool]$nativeSheet.Rows.Item(4).Hidden
    Write-Host ("  --> the hand-hidden row stayed hidden? {0}" -f $handHiddenStayedHidden)
    Write-Host '     (a TRUE here means the add-in must only reveal the rows its own'
    Write-Host '      relationships hid, and must not touch a hand-hidden row)'
    Write-Host

    Write-Host '=== VERDICT SUMMARY ==='
    Write-Host ('  Q1 scope:     the USER''s own group was collapsed too? {0}' -f $userGroupHidden)
    Write-Host '                 FALSE -> the mechanism is row-scoped, so the add-in'
    Write-Host '                          cannot touch a group it does not own.'
    Write-Host ('  Q2 refresh:   the group pass RE-EXPANDED the group? {0}' -f $groupPassReExpanded)
    Write-Host ('                 the reset pass RE-EXPANDED the group? {0}' -f $resetPassReExpanded)
    Write-Host '                 FALSE/FALSE -> a Refresh preserves the collapse, so the'
    Write-Host '                              Ribbon buttons MAY trigger a Refresh like the'
    Write-Host '                              other mutating commands do (R5.4 D11-amended).'
    Write-Host ('  Q3 restore:   every Gantt group expanded again? {0}' -f $allExpanded)
    Write-Host '  Q4 accessor:  `EntireRow.Hidden` is the write seam; a plain'
    Write-Host '                 `Range.Hidden` write is refused by the host.'
    Write-Host ('  Q5 identical: Excel''s own collapse HIDES the child rows? {0}' -f $nativeChildHidden)
    Write-Host ('                 its own expand SHOWS the child rows? {0}' -f $nativeChildShown)
    Write-Host '                 TRUE/TRUE -> our write produces exactly the state Excel'
    Write-Host '                              writes, so the Ribbon button is'
    Write-Host '                              indistinguishable from the native control.'
    Write-Host '  Candidate A (`Range.Outline.ShowLevels`) is REJECTED: `Range.Outline`'
    Write-Host '  is not exposed on IDispatch, and the worksheet-level call is'
    Write-Host '  sheet-wide, so it would collapse the user''s own groups.'

    foreach ($wb in @($excel.Workbooks)) { $wb.Close($false) }
}
finally {
    if ($excel) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    Write-Host
    Write-Host 'probe complete'
}
