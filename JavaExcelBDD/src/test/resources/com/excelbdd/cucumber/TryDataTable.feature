Feature: Try data table

    Scenario: Read table from feature file
        Given the following animals:
            | cow   |
            | horse |
            | sheep |

        And the cities basic status:
            | CityName | Popular | Mayor        |
            | New York | 2000    | Mark Tim     |
            | Shanghai | 2500    | Chen Liangyu |

        When key and value list:
            | header1 | header2 | header3 |
            | grid11  | grid12  | grid13  |
            | grid21  | grid22  | grid23  |