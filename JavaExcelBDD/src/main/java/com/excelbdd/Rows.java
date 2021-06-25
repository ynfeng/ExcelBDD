package com.excelbdd;

import java.util.List;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Rows {
    private int headerRow;
    private char parameterNameColumn;
    private String columnType;
    private final List<Row> rows = Lists.newArrayList();

    public void loadRows(XSSFSheet xssfSheet) {
        for (int iRow = 0; iRow < xssfSheet.getLastRowNum(); iRow++) {
            XSSFRow rowCurrent = xssfSheet.getRow(iRow);
            if (rowCurrent == null) {
                continue;
            }
            Row row = new Row(iRow, rowCurrent);
            rows.add(row);
            headerRow = row.headerRowNumber();
            parameterNameColumn = row.parameterNameColumnName();
            columnType = row.rowType();
            if (columnType != null) {
                break;
            }
        }
    }

    public int headerRowNumber() {
        return headerRow;
    }

    public char parameterNameColumnName() {
        return parameterNameColumn;
    }

    public String rowType() {
        return columnType;
    }
}
