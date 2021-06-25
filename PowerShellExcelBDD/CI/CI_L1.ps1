& $PSScriptRoot/InitializeTest.ps1

$TestPath = "$StartPath/PowerShellExcelBDD/Test"
$CoverageFile = "$StartPath/PowerShellExcelBDD/ExcelBDD/ExcelBDD.psm1"
$configuration = [PesterConfiguration]@{
    Run          = @{
        Path = $TestPath
        Exit = $false
    }
    Filter       = @{
        #Tag = 'Acceptance'
        #ExcludeTag = 'WindowsOnly'
    }
    Should       = @{
        ErrorAction = 'Continue'
    }
    CodeCoverage = @{
        Enabled = $true
        Path    = $CoverageFile
        # OutputFileFormat = 'CoverageGutters'
        OutputFileFormat = 'JaCoCo'
    }
    TestResult   = @{
        Enabled      = $true
        OutputFormat = 'JUnitXml'
    }
    Output       = @{
        Verbosity = 'Detailed'
    }
}
Invoke-Pester -Configuration $configuration -Verbose