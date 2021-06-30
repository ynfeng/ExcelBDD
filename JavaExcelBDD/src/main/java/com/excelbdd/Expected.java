package com.excelbdd;

public class Expected implements RowType {
    @Override
    public int actualParameterStartRow(int parameterRowNumber) {
        return parameterRowNumber + 1;
    }

    @Override
    public int columnStep() {
        return 2;
    }
}
