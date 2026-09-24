<#
.SYNOPSIS
    Installs the ProgressLogger macro into the Jupiter schedule workbook and saves it as .xlsm.

.DESCRIPTION
    - Imports ProgressLogger.bas (daily + weekly snapshot logic).
    - Adds the Workbook_Open / Workbook_BeforeSave events to ThisWorkbook.
    - Adds a "Record progress now" button on the Dashboard sheet.
    - Saves <workbook>.xlsm next to the source .xlsx (the .xlsx is left untouched).

    Easiest: double-click Install.cmd. Requires Windows + Microsoft Excel (2016 / 2019 / 2021 / 365).
    The workbook is looked up in this folder and in its parent folder unless -Workbook is given.
    -EnableVbomAccess switches on "Trust access to the VBA project object model" for this run only.

.EXAMPLE
    .\Install-ProgressLogger.ps1 -EnableVbomAccess
    .\Install-ProgressLogger.ps1 -Workbook "C:\Projects\Jupiter\Property_Scope_Plan_Jupiter_v2.xlsx" -EnableVbomAccess
#>
[CmdletBinding()]
param(
    [string]$Workbook,
    [switch]$EnableVbomAccess,
    [string]$OfficeVersion = '16.0'
)
$ErrorActionPreference = 'Stop'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$step = 'start'
function Step([string]$msg) { $script:step = $msg; Write-Host "-> $msg" }

$excel = $null; $wb = $null; $changedVbom = $false; $prevVbom = $null
$secKey = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security"
try {
    Step 'Locating files'
    $bas    = Join-Path $here 'ProgressLogger.bas'
    $events = Join-Path $here 'ThisWorkbook.txt'
    foreach ($f in @($bas, $events)) { if (-not (Test-Path -LiteralPath $f)) { throw "Missing file next to the installer: $f" } }
    try { Get-ChildItem -LiteralPath $here -File | Unblock-File -ErrorAction SilentlyContinue } catch { }

    if ($Workbook) {
        if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook" }
        $source = (Resolve-Path -LiteralPath $Workbook).Path
    } else {
        $parent = Split-Path -Parent $here
        $candidates = @()
        foreach ($dir in @($here, $parent)) {
            if ($dir) { $candidates += Get-ChildItem -LiteralPath $dir -Filter 'Property_Scope_Plan_Jupiter*.xlsx' -File -ErrorAction SilentlyContinue }
        }
        $candidates = @($candidates | Where-Object { $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending)
        if ($candidates.Count -eq 0) {
            throw "Could not find Property_Scope_Plan_Jupiter*.xlsx in '$here' or '$parent'. Put the workbook next to the installer or pass -Workbook <path>."
        }
        $source = $candidates[0].FullName
    }
    try { Unblock-File -LiteralPath $source -ErrorAction SilentlyContinue } catch { }
    $target = [IO.Path]::ChangeExtension($source, '.xlsm')
    Write-Host "   Workbook: $source"
    Write-Host "   Output  : $target"

    Step 'Checking Excel permission to access the VBA project'
    $policyKey = "HKCU:\Software\Policies\Microsoft\Office\$OfficeVersion\Excel\Security"
    $policy = (Get-ItemProperty -Path $policyKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
    if ($policy -eq 0) {
        throw "Your IT policy blocks programmatic access to VBA projects (AccessVBOM = 0 under Policies). Use the manual installation in README.md."
    }
    if (Test-Path $secKey) { $prevVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM }
    if ($prevVbom -ne 1) {
        if (-not $EnableVbomAccess) {
            throw ("Excel does not trust access to the VBA project object model. Re-run with -EnableVbomAccess (temporary) " +
                   "or enable it: File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'.")
        }
        if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
        New-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
        $changedVbom = $true
    }

    Step 'Starting Excel'
    if (Test-Path -LiteralPath $target) {
        try { [IO.File]::Open($target, 'Open', 'ReadWrite', 'None').Close() }
        catch { throw "The output file is open (close it in Excel first): $target" }
    }
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false          # do not fire Workbook_Open while installing

    Step 'Opening workbook'
    $wb = $excel.Workbooks.Open($source)
    try { $project = $wb.VBProject; $null = $project.VBComponents.Count }
    catch { throw "Excel refused access to the VBA project. Enable 'Trust access to the VBA project object model' (see README) and retry. Detail: $($_.Exception.Message)" }

    Step 'Importing ProgressLogger module'
    foreach ($c in @($project.VBComponents)) {
        if ($c.Name -eq 'ProgressLogger') { $project.VBComponents.Remove($c) }
    }
    $null = $project.VBComponents.Import($bas)

    Step 'Adding workbook events (ThisWorkbook)'
    $thisWbComp = $null
    if ($wb.CodeName) { try { $thisWbComp = $project.VBComponents.Item($wb.CodeName) } catch { } }
    if (-not $thisWbComp) {
        # Locale-independent fallback: the workbook document module exposes the workbook's own properties
        foreach ($c in @($project.VBComponents)) {
            if ($c.Type -ne 100) { continue }
            try { $null = $c.Properties.Item('FullName'); $thisWbComp = $c; break } catch { }
        }
    }
    if (-not $thisWbComp) { throw "Could not find the ThisWorkbook module. Use the manual installation (README), step 3." }
    $cm = $thisWbComp.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString([IO.File]::ReadAllText($events))

    Step 'Adding Dashboard button'
    try {
        $dash = $wb.Worksheets.Item('Dashboard')
        foreach ($s in @($dash.Shapes)) { if ($s.Name -eq 'btnRecordProgress') { $s.Delete() } }
        $anchor = $dash.Range('M3')
        $btn = $dash.Shapes.AddFormControl(0, $anchor.Left, $anchor.Top, 170, 26)   # 0 = xlButtonControl
        $btn.Name = 'btnRecordProgress'
        $btn.OnAction = 'RecordProgressNow'
        $btn.TextFrame.Characters().Text = 'Record progress now'
    } catch {
        Write-Warning "Button not added ($($_.Exception.Message)). The macro still works: Alt+F8 > RecordProgressNow."
    }

    Step 'Saving .xlsm'
    $wb.SaveAs($target, 52)               # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    $wb = $null
    Write-Host ''
    Write-Host "Installed. Open: $target" -ForegroundColor Green
    Write-Host "First open: click 'Enable Content' so the macro can record progress."
    $exitCode = 0
}
catch {
    Write-Host ''
    Write-Host "FAILED at step: $step" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    $exitCode = 1
}
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($changedVbom) {
        if ($null -eq $prevVbom) { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $prevVbom }
    }
}
exit $exitCode
