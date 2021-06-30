package com.excelbdd;

public interface RowType {
    RowType NULL = new RowType() {
        @Override
        public int actualParameterStartRow(int parameterRowNumber) {
            return -1;
        }

        @Override
        public int columnStep() {
            return 0;
        }
    };

    int actualParameterStartRow(int parameterRowNumber);

    int columnStep();
}
