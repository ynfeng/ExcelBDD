package com.excelbdd;

public class Simple implements RowType {
    @Override
    public int actualParameterStartRow(int parameterRowNumber) {
        return parameterRowNumber;
    }

    @Override
    public int columnStep() {
        return 1;
    }
}
