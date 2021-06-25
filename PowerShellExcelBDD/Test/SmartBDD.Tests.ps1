$script:ExcelBDDFilePath = "$StartPath/BDDExcel/ExcelBDD.xlsx"

Describe "Get BDD Data only by sheet" {

    $BDDTestCaseList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SmartBDD' 

    It "Get-SmartExampleList" -TestCases $BDDTestCaseList {
        Write-Host "Sheet $SheetName Column $Header"
        $TestExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        $TestcaseList = Get-ExampleList -ExcelPath $TestExcelPath `
            -WorksheetName $SheetName `
            -HeaderMatcher $HeaderMatcher `
            -HeaderUnmatcher $HeaderUnmatcher

        # Write-Host ($TestcaseList | ConvertTo-Json )
        
        $TestcaseList[0]["Header"] | Should -Be $Header1Name
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstGridValue
        $TestcaseList[1]["ParamName1"] | Should -Be $ParamName1InSet2Value
        $TestcaseList[1]["ParamName2"] | Should -Be $ParamName2InSet2Value
        $TestcaseList[0]["ParamName3"] | Should -Be $ParamName3Value
        $TestcaseList[0].Count | Should -Be $ParameterCount
        $MaxBlankThreshold | Should -Be 3
        $TestcaseList.Count | Should -Be $TestDataSetCount

        $TestcaseList[0]["ParamName3"] | Should -Be ""
        $TestcaseList[1]["ParamName3"] | Should -Be ""
        $TestcaseList[2]["ParamName3"] | Should -Be ""
        $TestcaseList[3]["ParamName3"] | Should -Be ""

        $TestcaseList[0]["ParamName4"] | Should -Be "2021/4/30"
        $TestcaseList[1]["ParamName4"] | Should -Be "0"
        $TestcaseList[2]["ParamName4"] | Should -Be "1"
        $TestcaseList[3]["ParamName4"] | Should -Be $LastGridValue
    }
}

Describe "Cover SBT" {

    $SpecificationByTestcaseList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'SpecificationByTestcase' 
    It "SpecificationByTestcase" -Testcases $SpecificationByTestcaseList {
        $Error.Clear()

        $Header | Should -Be $HeaderRowExpected
        ($HeaderRowTestResult -eq 'pass') | Should -Be $true
        $HeaderRowTestResult | Should -Be 'pass'
        Write-Host "SheetName:"$SheetName
        Write-Host "ParameterNameColumn:"$ParameterNameColumn
        Write-Host "HeaderRow:"$HeaderRow

        $TestExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        $TestcaseList = Get-ExampleList -ExcelPath $TestExcelPath `
            -WorksheetName $SheetName 
        
        $Error.Count | Should -Be 0

        [bool]$TestcaseList | Should -Be ($ExcelFileNameExpected -eq 'got')
        [bool]$TestcaseList | Should -Be ($SheetNameExpected -eq 'got')
        # Write-Host ($TestcaseList | ConvertTo-Json -Depth 10)

        $TestcaseList[0]["$FirstSetFirstCheckedParam"] | Should -Be $FirstSetFirstCheckedParamExpected
        $TestcaseList[0]["ParamName1"] | Should -Be $FirstSetParamName1
        $TestcaseList[0]["ParamName2"] | Should -Be $FirstSetParamName2
        $TestcaseList[0]["ParamName3"] | Should -Be $FirstSetParamName3
        $TestcaseList[0]["ParamName4"] | Should -Be $FirstSetParamName4

        $TestcaseList[0]["ParamName1Expected"] | Should -Be $FirstSetParamName1Expected
        $TestcaseList[0]["ParamName2Expected"] | Should -Be $FirstSetParamName2Expected
        $TestcaseList[0]["ParamName3Expected"] | Should -Be $FirstSetParamName3Expected
        $TestcaseList[0]["ParamName4Expected"] | Should -Be $FirstSetParamName4Expected

        if ($TestResultSwitch -eq 'On') {
            $TestcaseList[0]["ParamName1TestResult"] | Should -Be $FirstSetParamName1TestResult
            $TestcaseList[0]["ParamName2TestResult"] | Should -Be $FirstSetParamName2TestResult
            $TestcaseList[0]["ParamName3TestResult"] | Should -Be $FirstSetParamName3TestResult
            $TestcaseList[0]["ParamName4TestResult"] | Should -Be $FirstSetParamName4TestResult
        }
    }
}