Feature: Get Test data Set List
    test data is put in excel, so they should be get as a List

    Scenario Outline: Get Testcase list from a simple sheet
        Given The Excel file is "BDDExcel/ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        When Get the test dataset list
        Then Test dataset list which contains 5 sets is got
        And The No. 1 element of list Key "Header", Value is "Scenario1"

        Examples:
            | SheetName | HeaderRow | ParameterColumn |
            | Sheet1    | 1         | B               |
            | Sheet2    | 1         | C               |

