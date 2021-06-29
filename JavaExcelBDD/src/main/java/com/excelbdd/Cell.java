package com.excelbdd;

import org.apache.poi.xssf.usermodel.XSSFCell;

public class Cell {
    public int colIdx;
    private XSSFCell xssfCell;

    public Cell(int colIdx, XSSFCell xssfCell) {
        this.colIdx = colIdx;
        this.xssfCell = xssfCell;
    }

    boolean isParameterNameCell() {
        return value().contains("Parameter Name");
    }

    private String value() {
        return xssfCell.getStringCellValue();
    }

    public char parameterNameColumn() {
        return (char) (colIdx + 65);
    }

    public String stringValue() {
        return xssfCell.getStringCellValue();
    }

    public boolean isInputCell() {
        return stringValue().equals("Input");
    }

    public boolean isTestResultCell() {
        return stringValue().equals("Test Result");
    }
}
