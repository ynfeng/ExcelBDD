
Describe "DataTable" {
    $ExcelBDDFilePath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTableList = Get-ExampleList -ExcelPath $ExcelBDDFilePath `
        -WorksheetName 'DataTableBDD' 

    It "Check DataTable Reading" -TestCases $DataTableList {
        $ExcelPath = "$StartPath/BDDExcel/$ExcelFileName"
        $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName $SheetName `
            -HeaderRow $HeaderRow -StartColumn $StartColumn
        $DataTable1.Count | Should -Be $TestSetCount
        $DataTable1.GetType().Name | Should -Be 'Object[]'
        $DataTable1[0].GetType().Name | Should -Be 'HashTable'
        $DataTable1[0]["Header1"] | Should -Be $FirstGridValue
        $DataTable1[5]["Header8"] | Should -Be $LastGridValue
        $DataTable1[5].Count | Should -Be $ColumnCount
    }


}

Describe "Use ImportExcel Only" {
    It "Read From Excel" {
        $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
        $DataTable1 = Import-Excel -Path $ExcelPath -WorksheetName "DataTable1" -StartRow 2 -StartColumn 1
        $DataTable1.Count | Should -Be 6
        Write-Host $DataTable1[0]."Header1"
    }
}

Describe "Use ExcelBDD to get DataTable" {
    $ExcelPath = "$StartPath/BDDExcel/DataTableBDD.xlsx"
    $DataTable1 = Get-DataTable -ExcelPath $ExcelPath -WorksheetName DataTable1 -HeaderRow 2

    It "Use the DataTable" -Testcases $DataTable1 {
        Write-Host $Header1
        Write-Host $Header2
        $Header1 | Should -Not -BeNullOrEmpty
    }
}