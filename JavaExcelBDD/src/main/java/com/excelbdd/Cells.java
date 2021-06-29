package com.excelbdd;

import java.util.List;
import java.util.Optional;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Cells {
    private static final String SIMPLE = "SIMPLE";
    private static final String TESTRESULT = "TESTRESULT";
    private static final String EXPECTED = "EXPECTED";
    private final List<Cell> cells = Lists.newArrayList();
    private final int rowNum;

    public Cells(int rowNum) {
        this.rowNum = rowNum;
    }

    public Cell get(int index) {
        return cells.get(index);
    }

    private int getHeaderRowNumber(Cell parameterNameCell) {
        if (hasInputCell(parameterNameCell.colIdx)) {
            return rowNum;
        } else {
            return rowNum + 1;
        }
    }

    boolean hasInputCell(int colIdx) {
        return cells.get(colIdx + 1).isInputCell();
    }

    boolean hasTestResultCell(int colIdx) {
        return cells.get(colIdx + 3).isTestResultCell();
    }

    private String rowType(Cell parameterNameCell) {
        if (hasInputCell(parameterNameCell.colIdx) && hasTestResultCell(parameterNameCell.colIdx)) {
            return TESTRESULT;
        } else if (hasInputCell(parameterNameCell.colIdx)) {
            return EXPECTED;
        } else {
            return SIMPLE;
        }
    }

    public Optional<Cell> findParameterNameCell() {
        return cells.stream()
            .filter(Cell::isParameterNameCell)
            .findFirst();
    }

    public int headerRowNumber() {
        return findParameterNameCell().map(this::getHeaderRowNumber).orElse(-1);
    }

    public char parameterNameColumnName() {
        return findParameterNameCell().map(Cell::parameterNameColumn).orElse('@');
    }

    public String rowType() {
        return findParameterNameCell().map(this::rowType).orElse(null);
    }

    void loadCells(XSSFRow xssfRow) {
        for (int iCol = 0; iCol < xssfRow.getLastCellNum(); iCol++) {
            XSSFCell cellCurrent = xssfRow.getCell(iCol);
            if (cellCurrent == null) {
                continue;
            }
            cells.add(new Cell(iCol, cellCurrent));
        }
    }
}
