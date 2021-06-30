package com.excelbdd;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Sheet {
    private final XSSFSheet xssfSheet;
    public Rows rows = new Rows();

    public Sheet(XSSFSheet xssfSheet) {
        this.xssfSheet = xssfSheet;
        rows.loadRows(xssfSheet);
    }

    public void close() throws IOException {
        xssfSheet.getWorkbook().close();
    }

    public List<Map<String, String>> getExamples(String headerMatcher, String headerUnmatcher) throws IOException {
        return getExamples(headerRow(), headerMatcher, headerUnmatcher);
    }

    public List<Map<String, String>> getExamples(int parameterRowNumber, String headerMatcher, String headerUnmatcher) throws IOException {
        // poi get row from 0, so 1st headerRow is at 0
        // by default, actualHeaderRow is below
        RowType rowType = rowType();
        int actualHeaderRow = parameterRowNumber - 1;
        int actualParameterStartRow = rowType.actualParameterStartRow(parameterRowNumber);
        int columnStep = rowType.columnStep();

        ArrayList<Map<String, String>> listTestSet = new ArrayList<>();
        // poi get column from 0, so Column A's Num is 0, 65 is A's ASCII code
        int parameterNameColumnNum = parameterNameColumn() - 65;

        XSSFRow rowHeader = xssfSheet.getRow(actualHeaderRow);
        HashMap<Integer, Integer> mapTestSetHeader = getHeaderMap(headerMatcher, headerUnmatcher, listTestSet,
            parameterNameColumnNum, rowHeader, columnStep);

        // Get ParameterNames HashMap
        HashMap<Integer, String> mapParameterName = getParameterNameMap(actualParameterStartRow, parameterNameColumnNum,
            xssfSheet);

        for (Map.Entry<Integer, String> aParameterName : mapParameterName.entrySet()) {
            int iRow = aParameterName.getKey();
            String strParameterName = aParameterName.getValue();
            XSSFRow rowCurrent = xssfSheet.getRow(iRow);

            for (Map.Entry<Integer, Integer> entryHeader : mapTestSetHeader.entrySet()) {
                int iCol = entryHeader.getKey();
                Map<String, String> mapTestSet = listTestSet.get(entryHeader.getValue());
                putParameter(strParameterName, rowCurrent, mapTestSet, iCol);
                if (columnStep > 1) {
                    putParameter(strParameterName + "Expected", rowCurrent, mapTestSet, iCol + 1);
                    if (columnStep == 3) {
                        putParameter(strParameterName + "TestResult", rowCurrent, mapTestSet, iCol + 2);
                    }
                }
            }
        }
        return listTestSet;
    }

    private static HashMap<Integer, String> getParameterNameMap(int parameterStartRow, int parameterNameColumnNum, XSSFSheet sheetTestData) {
        HashMap<Integer, String> mapParameterName = new HashMap<>();
        int nContinuousBlankCount = 0;
        for (int iRow = parameterStartRow; iRow <= sheetTestData.getLastRowNum(); iRow++) {
            if (nContinuousBlankCount > 3) {
                break;
            }
            XSSFRow rowCurrent = sheetTestData.getRow(iRow);
            if (rowCurrent == null) {
                nContinuousBlankCount++;
                continue;
            }
            XSSFCell cellParameterName = rowCurrent.getCell(parameterNameColumnNum);
            if (cellParameterName == null) {
                nContinuousBlankCount++;
                continue;
            }
            String strParameterName = cellParameterName.getStringCellValue();
            if (strParameterName == null || strParameterName.isEmpty()) {
                nContinuousBlankCount++;
            } else if (strParameterName.equals("NA")) {
                nContinuousBlankCount = 0;
            } else {
                mapParameterName.put(iRow, strParameterName);
                nContinuousBlankCount = 0;
            }
        }
        return mapParameterName;
    }

    private static HashMap<Integer, Integer> getHeaderMap(String headerMatcher, String headerUnmatcher, ArrayList<Map<String, String>> listTestSet, int parameterNameColumnNum, XSSFRow rowHeader, int step) {
        // Get Matched Column HashMap
        String strRealHeaderMatcher = TestWizard.makeMatcherString(headerMatcher);
        String strRealHeaderUnmatcher;
        if (headerUnmatcher.isEmpty() || headerUnmatcher.equals(TestWizard.NEVER_MATCHED_STRING)) {
            strRealHeaderUnmatcher = TestWizard.NEVER_MATCHED_STRING;
        } else {
            strRealHeaderUnmatcher = TestWizard.makeMatcherString(headerUnmatcher);
        }
        int nMaxColumn = rowHeader.getLastCellNum();
        HashMap<Integer, Integer> mapTestSetHeader = new HashMap<>();
        int nTestSet = 0;
        for (int iCol = parameterNameColumnNum + 1; iCol < nMaxColumn; iCol += step) {
            XSSFCell cellHeader = rowHeader.getCell(iCol);
            String strHeader = cellHeader.getStringCellValue();
            if (strHeader != null && !strHeader.isEmpty() && strHeader.matches(strRealHeaderMatcher)
                && !strHeader.matches(strRealHeaderUnmatcher)) {
                mapTestSetHeader.put(iCol, nTestSet);
                Map<String, String> mapTestSet = new HashMap<>();
                mapTestSet.put("Header", cellHeader.getStringCellValue());
                listTestSet.add(mapTestSet);
                nTestSet++;
            }
        }
        return mapTestSetHeader;
    }

    private static void putParameter(String strParameterName, XSSFRow rowCurrent, Map<String, String> mapTestSet, int iCol) {

        XSSFCell cellCurrent = rowCurrent.getCell(iCol);
        if (cellCurrent.getCellType() == CellType.STRING) {
            mapTestSet.put(strParameterName, cellCurrent.getStringCellValue());
        } else if (cellCurrent.getCellType() == CellType.NUMERIC) {
            mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getNumericCellValue()));
        } else if (cellCurrent.getCellType() == CellType._NONE) {
            mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getDateCellValue()));
        } else if (cellCurrent.getCellType() == CellType.BLANK) {
            mapTestSet.put(strParameterName, "");
        } else if (cellCurrent.getCellType() == CellType.BOOLEAN) {
            mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getBooleanCellValue()));
        } else if (cellCurrent.getCellType() == CellType.FORMULA) {
            mapTestSet.put(strParameterName, cellCurrent.getRawValue());
        } else {
            mapTestSet.put(strParameterName, cellCurrent.getRawValue());
        }
    }

    public int headerRow() {
        return rows.headerRowNumber();
    }

    public char parameterNameColumn() {
        return rows.parameterNameColumnName();
    }

    public RowType rowType() {
        return rows.rowType();
    }
}
