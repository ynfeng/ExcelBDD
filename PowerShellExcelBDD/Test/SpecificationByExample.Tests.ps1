$script:ExcelBDDFilePath = "$StartPath/BDDExcel/ExcelBDD.xlsx"

Describe "Get BDD Data" {

    $BDDTestCaseList = Get-ExampleListByHeader -ExcelPath $ExcelBDDFilePath `
        -ParameterNameColumn F `
        -HeaderRow 1 `
        -HeaderMatcher Scenario

    # -WorksheetName 'SpecificationByExample' `

    It "Easy Success of Column List" -TestCases $BDDTestCaseList {
        Write-Host "Easy Success of Sheet $SheetName Column $Header"
        Write-Host "Header Row $HeaderRow"
        Write-Host "ParameterColumn $ParameterNameColumn"
        Write-Host "SheetName $SheetName"

        $TestcaseList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
            -WorksheetName $SheetName `
            -HeaderMatcher $HeaderMatcher `
            -HeaderUnmatcher $HeaderUnmatcher 

        Write-Host ($TestcaseList | ConvertTo-Json )
        
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

Describe "Test filter the dashboard by department" {
    $ExampleList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'StoryExample1' 
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