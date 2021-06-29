package com.excelbdd;

import java.util.List;
import java.util.Objects;
import org.apache.commons.compress.utils.Lists;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Rows {
    private final List<Row> rows = Lists.newArrayList();

    public void loadRows(XSSFSheet xssfSheet) {
        for (int iRow = 0; iRow < xssfSheet.getLastRowNum(); iRow++) {
            XSSFRow rowCurrent = xssfSheet.getRow(iRow);
            if (rowCurrent == null) {
                continue;
            }
            Row row = new Row(iRow, rowCurrent);
            rows.add(row);
        }
    }

    public int headerRowNumber() {
        return rows.stream()
            .map(Row::headerRowNumber)
            .filter(rowNumber -> rowNumber != -1)
            .findAny()
            .orElseThrow(() -> new IllegalStateException("header not found."));
    }

    public char parameterNameColumnName() {
        return rows.stream()
            .map(Row::parameterNameColumnName)
            .filter(columnName -> columnName != '@')
            .findAny()
            .orElseThrow(() -> new IllegalStateException("parameter name column not found."));
    }

    public String rowType() {
        return rows.stream()
            .map(Row::rowType)
            .filter(Objects::nonNull)
            .findAny()
            .orElseThrow(() -> new IllegalStateException("has not row type."));
    }
}
