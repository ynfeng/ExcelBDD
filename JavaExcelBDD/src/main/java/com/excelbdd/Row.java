package com.excelbdd;

import org.apache.poi.xssf.usermodel.XSSFRow;

public class Row {
    private final Cells cells;

    public Row(int rowIdx, XSSFRow xssfRow) {
        cells = new Cells(rowIdx);
        cells.loadCells(xssfRow);
    }

    public int headerRowNumber() {
        return cells.headerRowNumber();
    }

    public char parameterNameColumnName() {
        return cells.parameterNameColumnName();
    }

    public RowType rowType() {
        return cells.rowType();
    }
}
