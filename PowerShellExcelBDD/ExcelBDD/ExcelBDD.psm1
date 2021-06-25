#Define max blank lines for ending sheet reading
$MaxBlankThreshold = 3

<#
.Description
Get worksheet from Excel file according to build sheet path and worksheet name
.Example
Get-ExcelWorksheet -ExcelPath C:\buildsheet.xlsx -WorksheetName 'PaaS SQL DB Build'
#>
function Get-ExcelWorksheet {
    param (
        [String]$ExcelPath,
        [String]$WorksheetName
    )
    if (-not (Test-Path $ExcelPath)) {
        throw "$ExcelPath file does not exist."
    }
    try {
        $script:appExcel = Open-ExcelPackage -Path $ExcelPath
        if ($WorksheetName) {
            $Worksheet = $appExcel.Workbook.Worksheets[$WorksheetName]
        }
        else {
            $Worksheet = $appExcel.Workbook.Worksheets | Select-Object -First 1
        }
    }
    catch {
        $script:appExcel = New-Object -ComObject Excel.Application
        # Let Excel run in the backend, comment out below line, if debug, remove below #
        # $script:appExcel.Visible = $true
        $WorkBook = $script:appExcel.Workbooks.Open($ExcelPath)
        if ($WorksheetName) {
            $Worksheet = $WorkBook.Sheets[$WorksheetName]
        }
        else {
            $Worksheet = $WorkBook.Sheets[0]
        }
    }
    if ($null -eq $Worksheet ) {
        throw "$WorksheetName sheet does not exist."
    }
    return $Worksheet
}

function Close-ExcelWorksheet {
    try {
        if ($script:appExcel.Name -eq "Microsoft Excel") {
            $script:appExcel.ActiveWorkbook.Close($false)
            $script:appExcel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($script:appExcel)
        }
        else {
            Close-ExcelPackage -ExcelPackage $script:appExcel -NoSave
        }
    }
    catch {
        Write-Debug "Excel is closed."
    }
}

<#
.Description
Get a Hashtable list from excel sheet, one row for one hashtable
.Example
    #Get TestcaseDataList for Pester Testcase
    $TestcaseDataList = Get-DataTable -WorksheetName SheetName `
        -ExcelPath "${StartPath}${SEP}IaCSQLDBToolKit${SEP}TestData${SEP}DBTestCaseData.xlsx"  `
        -HeaderRow 1
    It "Full Rule Except Email From Excel File" -Testcases $TestcaseDataList {
        Test-MZIsPropertyValid -PropertyName $PropertyName -PropertyValue $PropertyValue -Rule $Rule | Should -Be ($Expected -eq "TRUE")
    }
#>
function Get-DataTable {
    param (
        [string]$ExcelPath,
        [string]$WorksheetName,
        [int]$HeaderRow = 1,
        [string]$StartColumn = 'A'
    )
    $IntStartColumn = [int][char]($StartColumn.ToUpper()) - 64
    $RawDataTableA = Import-Excel -Path $ExcelPath -WorksheetName $WorksheetName `
        -StartRow $HeaderRow -StartColumn $IntStartColumn
    $DataTableA = @()
    foreach ($item in $RawDataTableA) {
        $HashTableA = @{}
        $item.psobject.properties | ForEach-Object { $HashTableA[$_.Name] = $_.Value }
        $DataTableA += $HashTableA
    }
    return $DataTableA
}

<#
.Description
Get hashtable list of Example data, one Hashtable from one example data area in excel sheet
alias is Get-TestcaseList
.Example
Describe "Test Get-ExampleList" {
    use default HeaderRow which is 1, and default ParameterNameColumn which is C
    $ExampleList = Get-ExampleListByHeader -ExcelPath ".\Excel\Example1.xlsx" -WorksheetName 'Scenario1'
    It "Easy Success of SBE" -TestCases $ExampleList {
        [int]$BlackSweaterCountAtCustomer | Should -BeGreaterOrEqual $BlackSweaterCountReturned
        [int]$BlackSweaterCountInInvertory1 + [int]$BlackSweaterCountReturned | Should -Be $BlackSweaterCountInInvertory2
    }
}

Describe "Test filter the dashboard by department" {
    $ExampleList = Get-ExampleListByHeader -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'StoryExample1' `
        -ParameterNameColumn E `
        -HeaderRow 3
    It "Run Example one by one" -TestCases $ExampleList {
        #The below variables are generated automatically from Excel
        Write-Host "===$Header==="
        Write-Host $SelectedView
        Write-Host $DepartmentCount
        Write-Host $SelectedDepartment
        Write-Host $FullDepartmentName
        Write-Host $DepartmentLocation
        Write-Host $DepartmentCurrentMonthKPI1
        Write-Host $DepartmentCurrentMonthKPI2
    }
}
#>
function Get-ExampleListByHeader {
    param (
        [string]$ExcelPath,
        [string]$WorksheetName,
        [int]$HeaderRow = 1,
        [string]$ParameterNameColumn = 'C',
        [string]$HeaderMatcher,
        [string]$HeaderUnmatcher,
        [switch]$Expected,
        [switch]$TestResult
    )
    $Worksheet = Get-ExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName

    if ($HeaderRow.GetType().Name -eq 'String') {
        $HeaderRow = [int]$HeaderRow
    }

    return Get-ExampleListFromWorksheet -Worksheet $Worksheet `
        -HeaderRow $HeaderRow `
        -ParameterNameColumn $ParameterNameColumn `
        -HeaderMatcher $HeaderMatcher `
        -HeaderUnmatcher $HeaderUnmatcher `
        -Expected:$Expected `
        -TestResult:$TestResult
}


function Get-ExampleListFromWorksheet {
    param (
        $Worksheet,
        [int]$HeaderRow,
        [string]$ParameterNameColumn,
        [string]$HeaderMatcher,
        [string]$HeaderUnmatcher,
        [switch]$Expected,
        [switch]$TestResult
    )

    if ($TestResult) {
        $ColumnStep = 3
        $CurrentRow = $HeaderRow + 2
    }
    elseif ($Expected) {
        $ColumnStep = 2
        $CurrentRow = $HeaderRow + 2
    }
    else {
        $ColumnStep = 1
        $CurrentRow = $HeaderRow + 1
    }
    $ParamNameCol = [int][char]($ParameterNameColumn.ToUpper()) - 64
    #Get Test data set Column Array
    $CurrentCol = $ParamNameCol + 1
    $ColumnNumArray = @()
    while (-not [String]::IsNullOrEmpty($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text)) {
        if ($HeaderMatcher -and (-not $HeaderUnmatcher)) {
            if ($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text -match $HeaderMatcher) {
                $ColumnNumArray += $CurrentCol
            }
        }
        elseif ((-not $HeaderMatcher) -and $HeaderUnmatcher) {
            if ($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text -notmatch $HeaderUnmatcher) {
                $ColumnNumArray += $CurrentCol
            }
        }
        elseif ($HeaderMatcher -and $HeaderUnmatcher) {
            if (($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text -match $HeaderMatcher) `
                    -and ($Worksheet.Cells.Item($HeaderRow, $CurrentCol).Text -notmatch $HeaderUnmatcher)) {
                $ColumnNumArray += $CurrentCol
            }
        }
        else {
            $ColumnNumArray += $CurrentCol
        }
        $CurrentCol += $ColumnStep
    }

    #Get Parameter Row Array
    $RowNumArray = @()

    $ContinuousBlankCount = 0
    do {
        if ([String]::IsNullOrEmpty($Worksheet.Cells.Item($CurrentRow, $ParamNameCol).Text)) {
            $ContinuousBlankCount++
        }
        else {
            $ContinuousBlankCount = 0
            if ("NA" -ne $Worksheet.Cells.Item($CurrentRow, $ParamNameCol).Text) {
                $RowNumArray += $CurrentRow
            }
        }
        $CurrentRow++
    }while ($ContinuousBlankCount -le $MaxBlankThreshold)

    $List = [System.Collections.ArrayList]::new()
    foreach ($iCol in $ColumnNumArray) {
        $DataSet = [ordered]@{}
        #Put Header
        $DataSet["Header"] = $Worksheet.Cells.Item($HeaderRow, $iCol).Text.Trim()
        foreach ($iRow in $RowNumArray) {
            $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim()] = $Worksheet.Cells.Item($iRow, $iCol).Text
            if ($TestResult -or $Expected) {
                $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim() + "Expected"] = $Worksheet.Cells.Item($iRow, $iCol + 1).Text
                if ($TestResult) {
                    $DataSet[$Worksheet.Cells.Item($iRow, $ParamNameCol).Text.Trim() + "TestResult"] = $Worksheet.Cells.Item($iRow, $iCol + 2).Text
                }
            }
        }
        [void]$List.Add($DataSet)
    }
    Close-ExcelWorksheet | Out-Null
    return $List
}


<#
.Description
Get hashtable list of Example data, one Hashtable from one example data area in excel sheet

.Example
Describe "Test Get-ExampleList" {
    $ExampleList = Get-ExampleList -ExcelPath ".\Excel\Example1.xlsx" -WorksheetName 'Scenario1'
    It "Easy Success of SBE" -TestCases $ExampleList {
        [int]$BlackSweaterCountAtCustomer | Should -BeGreaterOrEqual $BlackSweaterCountReturned
        [int]$BlackSweaterCountInInvertory1 + [int]$BlackSweaterCountReturned | Should -Be $BlackSweaterCountInInvertory2
    }
}

Describe "Test filter the dashboard by department" {
    $TestcaseList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'StoryExample1'
    It "Run Example one by one" -TestCases $TestcaseList {
        #The below variables are generated automatically from Excel
        Write-Host "===$Header==="
        Write-Host $SelectedView
        Write-Host $DepartmentCount
        Write-Host $SelectedDepartment
        Write-Host $FullDepartmentName
        Write-Host $DepartmentLocation
        Write-Host $DepartmentCurrentMonthKPI1
        Write-Host $DepartmentCurrentMonthKPI2
    }
}
#>
function Get-ExampleList {
    param (
        [string]$ExcelPath,
        [string]$WorksheetName,
        [string]$HeaderMatcher,
        [string]$HeaderUnmatcher
    )
    $Worksheet = Get-ExcelWorksheet -ExcelPath $ExcelPath -WorksheetName $WorksheetName
    for ($iRow = 1; $iRow -le $Worksheet.Dimension.Rows; $iRow++) {
        for ($iColumn = 1; $iColumn -lt $Worksheet.Dimension.Columns; $iColumn++) {
            if ($Worksheet.Cells.Item($iRow, $iColumn).Text -match "^Param.*Name") {
                # [int][char]($ParameterNameColumn.ToUpper()) - 64
                $ParameterNameColumn = [string][char]($iColumn + 64)
                if ($Worksheet.Cells.Item($iRow, $iColumn + 1).Text -match 'Input') {
                    $HeaderRow = $iRow - 1
                    if ($Worksheet.Cells.Item($iRow, $iColumn + 3).Text -match 'Test Result') {
                        $TestResult = $true
                    }
                    else {
                        $Expected = $true
                    }
                }
                else {
                    $HeaderRow = $iRow
                }
                Break
            }
        }
        if ($HeaderRow) {
            Break
        }
    }
    if (-Not $HeaderRow) {
        throw "Parameter Name grid is not found"
    }
    return Get-ExampleListFromWorksheet -Worksheet $Worksheet `
        -HeaderRow $HeaderRow `
        -ParameterNameColumn $ParameterNameColumn `
        -HeaderMatcher $HeaderMatcher `
        -HeaderUnmatcher $HeaderUnmatcher `
        -Expected:$Expected `
        -TestResult:$TestResult
}