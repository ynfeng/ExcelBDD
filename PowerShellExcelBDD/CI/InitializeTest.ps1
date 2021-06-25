Get-Module ExcelBDD | Remove-Module
Import-Module Pester
Import-Module ImportExcel
$global:StartPath = (Resolve-Path "$PSScriptRoot/../..").Path
Write-Host $global:StartPath
$modulePath = Join-Path $StartPath "PowerShellExcelBDD/ExcelBDD/ExcelBDD.psm1"
Import-Module $modulePath