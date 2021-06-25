Feature: Get Test data Set List from Excel Specification by Testcase
    Specification and its testcase are put in excel, so they should be get as a List

    Scenario Outline: Get list from a simple sheet
        Given The Excel file is "BDDExcel/ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"

        When invoke get test data from excel
        Then a testset list is got, which count is <ListCount>
        And The Header of No. 1 set is "Scenario1"
        And Input value of Variable "ParamName1" of No. 1 set is "V1.1"
        And Expected value of Variable "ParamName1" of No. 1 set is "V1.1"
        And Expected value of Variable "ParamName1" of No. 2 set is "V1.2"
        And Expected value of Variable "ParamName1" of No. 3 set is "V1.3"
        And Expected value of Variable "ParamName1" of No. 4 set is "V1.4"
        And Test Result value of Variable "ParamName1" of No. 1 set is "pass"
        And The 1st data table is:
            | ParameterName | Input     | Expected  | TestResult |
            | ParamName1    | V1.1      | V1.1      | pass       |
            | ParamName2    | V2.1      | V2.1      | pass       |
            | ParamName3    |           |           | pass       |
            | ParamName4    | 2021/4/30 | 2021/4/30 | pass       |

        Examples:
            | SheetName | HeaderRow | ParameterColumn | ListCount |
            | SBTSheet1 | 1         | B               | 5         |
            | SBTSheet3 | 1         | D               | 6         |


    Scenario Outline: Get list according to header matching
        Given The Excel file is "BDDExcel/ExcelBDD.xlsx"
        And The Sheet name is "<SheetName>"
        And Header Row is <HeaderRow>
        And Parameter Column is "<ParameterColumn>"
        And Matcher is "<Matcher>"
        When invoke get test data from excel according to Matcher
        Then a testset list is got, which count is <ListCount>
        And The Header of No. 1 set is "<HeaderName>"
        And Input value of Variable "ParamName1" of No. 1 set is "V1.1"
        And Expected value of Variable "ParamName1" of No. 1 set is "V1.1"
        And Test Result value of Variable "ParamName1" of No. 1 set is "pass"
        And The 1st data table is:
            | ParameterName | Input     | Expected  | TestResult |
            | ParamName1    | V1.1      | V1.1      | pass       |
            | ParamName2    | V2.1      | V2.1      | pass       |
            | ParamName3    |           |           | pass       |
            | ParamName4    | 2021/4/30 | 2021/4/30 | pass       |
        And Expected value of Variable "ParamName1" of No. <ListCount> set is "<ParamName1Value>"

        Examples:
            | SheetName | HeaderRow | ParameterColumn | Matcher    | ListCount | ParamName1Value | HeaderName |
            | SBTSheet3 | 1         | D               | Scenario   | 6         | V1.1            | Scenario1  |
            | SBTSheet3 | 1         | D               | Scenario1  | 2         | V1.1            | Scenario1  |
            | SBTSheet3 | 1         | D               | Scenario1b | 1         | V1.1            | Scenario1b |