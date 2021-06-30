package com.excelbdd;

public class TestResult implements RowType {
    @Override
    public int actualParameterStartRow(int parameterRowNumber) {
        return parameterRowNumber + 1;
    }

    @Override
    public int columnStep() {
        return 3;
    }
}
