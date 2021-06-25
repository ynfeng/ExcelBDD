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

    private int getHeaderRowNumber(Cell cell) {
        if (cell.hasInputCell(this)) {
            return rowNum;
        } else {
            return rowNum + 1;
        }
    }

    private String rowType(Cell cell) {
        if (cell.hasInputCell(this) && cell.hasTestResultCell(this)) {
            return TESTRESULT;
        } else if (cell.hasInputCell(this)) {
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
