<#
.SYNOPSIS
  Fallback macro packaging with Excel: open an .xlsx built by fepm, import vba/*.bas, optionally run
  RunAllScenarios, and save as .xlsm (FileFormat 52).
.NOTES
  Needs desktop Excel and "Trust access to the VBA project object model"
  (File > Options > Trust Center > Trust Center Settings > Macro Settings). Without it $wb.VBProject throws.
.EXAMPLE
  powershell -ExecutionPolicy Bypass -File tools\inject_vba.ps1 -Workbook projects\_demo\out\Demo_Wind_Bid.xlsx -Run
#>
param(
  [Parameter(Mandatory = $true)][string]$Workbook,
  [string]$VbaDir = (Join-Path $PSScriptRoot "..\vba"),
  [string]$Out,
  [switch]$Run
)
$ErrorActionPreference = "Stop"
$src = (Resolve-Path $Workbook).Path
if (-not $Out) { $Out = [IO.Path]::ChangeExtension($src, ".xlsm") }
$Out = [IO.Path]::GetFullPath($Out)
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
  $wb = $excel.Workbooks.Open($src)
  $proj = $wb.VBProject
  foreach ($c in @($proj.VBComponents)) { if ($c.Type -eq 1) { $proj.VBComponents.Remove($c) } }  # 1 = standard module
  Get-ChildItem (Resolve-Path $VbaDir) -Filter *.bas | ForEach-Object { [void]$proj.VBComponents.Import($_.FullName) }
  if ($Run) { $excel.Run("RunAllScenarios") }
  $wb.SaveAs($Out, 52)
  $wb.Close($false)
  Write-Host "saved $Out"
} finally {
  $excel.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
