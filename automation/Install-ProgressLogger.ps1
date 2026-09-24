<#
.SYNOPSIS
    Installs the ProgressLogger macro into the Jupiter schedule workbook and saves it as .xlsm.

.DESCRIPTION
    - Imports ProgressLogger.bas (daily + weekly snapshot logic).
    - Adds the Workbook_Open / Workbook_BeforeSave events to ThisWorkbook.
    - Adds a "Record progress now" button on the Dashboard sheet.
    - Saves <workbook>.xlsm next to the source .xlsx (the .xlsx is left untouched).

    Requires Windows + Microsoft Excel (2016 / 2019 / 2021 / Microsoft 365).
    Excel must allow programmatic access to the VBA project. Pass -EnableVbomAccess to switch it on
    only for this run (the previous setting is restored at the end).

.EXAMPLE
    .\Install-ProgressLogger.ps1 -Workbook ..\Property_Scope_Plan_Jupiter_v2.xlsx -EnableVbomAccess
#>
[CmdletBinding()]
param(
    [string]$Workbook = (Join-Path $PSScriptRoot '..\Property_Scope_Plan_Jupiter_v2.xlsx'),
    [switch]$EnableVbomAccess,
    [string]$OfficeVersion = '16.0'
)
$ErrorActionPreference = 'Stop'

$source = (Resolve-Path $Workbook).Path
$target = [IO.Path]::ChangeExtension($source, '.xlsm')
$bas    = Join-Path $PSScriptRoot 'ProgressLogger.bas'
$events = Join-Path $PSScriptRoot 'ThisWorkbook.txt'
foreach ($f in @($bas, $events)) { if (-not (Test-Path $f)) { throw "Missing file: $f" } }

# Programmatic access to the VBA project (File > Options > Trust Center > Macro Settings)
$secKey  = "HKCU:\Software\Microsoft\Office\$OfficeVersion\Excel\Security"
$prevVbom = $null
if (Test-Path $secKey) { $prevVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM }
$changedVbom = $false
if ($prevVbom -ne 1) {
    if (-not $EnableVbomAccess) {
        throw ("Excel does not trust access to the VBA project object model. Re-run with -EnableVbomAccess " +
               "(temporary, restored at the end) or enable it in Excel: File > Options > Trust Center > " +
               "Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'.")
    }
    if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
    New-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force | Out-Null
    $changedVbom = $true
}

$excel = $null
$wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false          # do not fire Workbook_Open while installing

    $wb = $excel.Workbooks.Open($source)
    $project = $wb.VBProject

    # Replace a previous install, if any
    foreach ($c in @($project.VBComponents)) {
        if ($c.Name -eq 'ProgressLogger') { $project.VBComponents.Remove($c) }
    }
    $null = $project.VBComponents.Import($bas)

    # ThisWorkbook module name is localized (e.g. "EsteLibro"); CodeName always points to it
    $thisWb = $project.VBComponents.Item($wb.CodeName).CodeModule
    if ($thisWb.CountOfLines -gt 0) { $thisWb.DeleteLines(1, $thisWb.CountOfLines) }
    $thisWb.AddFromString((Get-Content $events -Raw))

    # Dashboard button
    $dash = $wb.Worksheets.Item('Dashboard')
    foreach ($s in @($dash.Shapes)) { if ($s.Name -eq 'btnRecordProgress') { $s.Delete() } }
    $anchor = $dash.Range('M3')
    $btn = $dash.Shapes.AddFormControl(0, $anchor.Left, $anchor.Top, 170, 26)   # 0 = xlButtonControl
    $btn.Name = 'btnRecordProgress'
    $btn.OnAction = 'RecordProgressNow'
    $btn.TextFrame.Characters().Text = 'Record progress now'

    if (Test-Path $target) { Remove-Item $target -Force }
    $wb.SaveAs($target, 52)               # 52 = xlOpenXMLWorkbookMacroEnabled
    $wb.Close($false)
    $wb = $null
    Write-Host "Installed. Open: $target" -ForegroundColor Green
    Write-Host "First open: click 'Enable Content' so the macro can record progress."
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($changedVbom) {
        if ($null -eq $prevVbom) { Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $prevVbom }
    }
}
