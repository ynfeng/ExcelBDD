package com.excelbdd;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;

import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Behavior {
	private static final String SIMPLE = "SIMPLE";
	private static final String TESTRESULT = "TESTRESULT";
	private static final String EXPECTED = "EXPECTED";

	private Behavior() {
	}

	public static List<Map<String, String>> getExampleList(String excelPath) throws IOException {
		return getExampleList(excelPath, "", TestWizard.ANY_MATCHER);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath) throws IOException {
		return getExampleList(excelPath).stream();
	}

	/**
	 * @param excelPath
	 * @param worksheetName
	 * @return Example List
	 * @throws IOException
	 */
	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName) throws IOException {
		return getExampleList(excelPath, worksheetName, TestWizard.ANY_MATCHER);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName)
			throws IOException {
		return getExampleList(excelPath, worksheetName).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, String headerMatcher)
			throws IOException {
		return getExampleList(excelPath, worksheetName, headerMatcher, TestWizard.NEVER_MATCHED_STRING);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName,
			String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerMatcher).stream();
	}

	@SuppressWarnings("resource")
	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, String headerMatcher,
			String headerUnmatcher) throws IOException {
		// Find the Header Row and Parameter Name Column

		FileInputStream excelFile = new FileInputStream(new File(excelPath));
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		XSSFSheet sheetTestData = getExampleSheet(worksheetName, excelFile, workbook);

		int headerRow = 0;
		char parameterNameColumn = 0;
		String columnType = null;
		for (int iRow = 0; iRow < sheetTestData.getLastRowNum(); iRow++) {
			XSSFRow rowCurrent = sheetTestData.getRow(iRow);
			if (rowCurrent == null) {
				continue;
			}
			for (int iCol = 0; iCol < rowCurrent.getLastCellNum(); iCol++) {
				XSSFCell cellCurrent = rowCurrent.getCell(iCol);
				if (cellCurrent == null) {
					continue;
				}
				String cellValue;
				if (cellCurrent.getCellType().equals(CellType.STRING)) {
					cellValue = cellCurrent.getStringCellValue();
				} else {
					continue;
				}
				if (isParameterNameGrid(cellValue)) {
					parameterNameColumn = (char) (iCol + 65);
					if (hasInputGrid(rowCurrent, iCol)) {
						headerRow = iRow;
						if (hasTestResultGrid(rowCurrent, iCol)) {
							columnType = TESTRESULT;
						} else {
							columnType = EXPECTED;
						}
					} else {
						columnType = SIMPLE;
						headerRow = iRow + 1;
					}
					break;
				}
			}
			if (columnType != null) {
				break;
			}
		}
		if (columnType == null) {
			throw new IOException("Parameter Name grid is not found.");
		}
		return getExampleListFromWorksheet(excelFile, sheetTestData, headerRow, parameterNameColumn, headerMatcher,
				headerUnmatcher, columnType);
	}

	protected static boolean hasInputGrid(XSSFRow rowCurrent, int iCol) {
		try {
			return rowCurrent.getCell(iCol + 1).getStringCellValue().equals("Input");
		} catch (NullPointerException e) {
			return false;
		}
	}

	protected static boolean hasTestResultGrid(XSSFRow rowCurrent, int iCol) {
		try {
			return rowCurrent.getCell(iCol + 3).getStringCellValue().equals("Test Result");
		} catch (NullPointerException e) {
			return false;
		}
	}

	protected static XSSFSheet getExampleSheet(String worksheetName, FileInputStream excelFile, XSSFWorkbook workbook)
			throws IOException {
		XSSFSheet sheetTestData;
		if (worksheetName.isEmpty()) {
			sheetTestData = workbook.getSheetAt(0);
		} else {
			sheetTestData = workbook.getSheet(worksheetName);
		}
		if (sheetTestData == null) {
			workbook.close();
			excelFile.close();
			throw new IOException(worksheetName + " sheet does not exist.");
		}
		return sheetTestData;
	}

	private static boolean isParameterNameGrid(String cellValue) {
		return cellValue.matches("Param.*Name.*");
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName,
			String headerMatcher, String headerUnmatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerMatcher, headerUnmatcher).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, TestWizard.ANY_MATCHER,
				TestWizard.NEVER_MATCHED_STRING, SIMPLE);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, TestWizard.ANY_MATCHER,
				TestWizard.NEVER_MATCHED_STRING, SIMPLE).stream();
	}

	/**
	 * @param excelPath
	 * @param worksheetName
	 * @param headerRow
	 * @param parameterNameColumn
	 * @param headerMatcher
	 * @return Example list
	 * @throws IOException
	 */
	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				TestWizard.NEVER_MATCHED_STRING);
	}

	/**
	 * @param excelPath
	 * @param worksheetName
	 * @param headerRow
	 * @param parameterNameColumn
	 * @param headerMatcher
	 * @return Example Stream
	 * @throws IOException
	 */
	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher).stream();
	}

	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnmatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnmatcher,
				SIMPLE);
	}

	public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnmatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnmatcher)
				.stream();
	}

	private static HashMap<Integer, Integer> getHeaderMap(String headerMatcher, String headerUnmatcher,
			ArrayList<Map<String, String>> listTestSet, int parameterNameColumnNum, XSSFRow rowHeader, int step) {
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
			if (cellHeader == null) {
				break;
			}
			String strHeader = cellHeader.getStringCellValue();
			if (isHeaderValid(strRealHeaderMatcher, strRealHeaderUnmatcher, strHeader)) {
				mapTestSetHeader.put(iCol, nTestSet);
				Map<String, String> mapTestSet = new HashMap<>();
				mapTestSet.put("Header", cellHeader.getStringCellValue());
				listTestSet.add(mapTestSet);
				nTestSet++;
			}
		}
		return mapTestSetHeader;
	}

	private static boolean isHeaderValid(String strRealHeaderMatcher, String strRealHeaderUnmatcher, String strHeader) {
		return (strHeader != null) && (!strHeader.isEmpty()) && strHeader.matches(strRealHeaderMatcher)
				&& (!strHeader.matches(strRealHeaderUnmatcher));
	}

	private static HashMap<Integer, String> getParameterNameMap(int parameterStartRow, int parameterNameColumnNum,
			XSSFSheet sheetTestData) {
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

	public static Collection<Object[]> getExampleCollection(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn) throws IOException {
		Collection<Object[]> collectionTestData = new ArrayList<>();
		List<Map<String, String>> listTestData = getExampleList(excelPath, worksheetName, headerRow,
				parameterNameColumn, TestWizard.ANY_MATCHER, TestWizard.NEVER_MATCHED_STRING, SIMPLE);
		for (Map<String, String> map : listTestData) {
			Object[] arrayObj = { map };
			collectionTestData.add(arrayObj);
		}
		return collectionTestData;
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, TestWizard.ANY_MATCHER,
				TestWizard.NEVER_MATCHED_STRING, EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				TestWizard.NEVER_MATCHED_STRING, EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher, String headerUnmatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnmatcher,
				EXPECTED);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn) throws IOException {
		return getExampleListWithTestResult(excelPath, worksheetName, headerRow, parameterNameColumn,
				TestWizard.ANY_MATCHER);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher,
				TestWizard.NEVER_MATCHED_STRING, TESTRESULT);
	}

	public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName,
			int headerRow, char parameterNameColumn, String headerMatcher, String headerUnmatcher) throws IOException {
		return getExampleList(excelPath, worksheetName, headerRow, parameterNameColumn, headerMatcher, headerUnmatcher,
				TESTRESULT);
	}

	@SuppressWarnings("resource")
	public static List<Map<String, String>> getExampleList(String excelPath, String worksheetName, int headerRow,
			char parameterNameColumn, String headerMatcher, String headerUnmatcher, String columnType)
			throws IOException {

		FileInputStream excelFile = new FileInputStream(new File(excelPath));
		XSSFWorkbook workbook = new XSSFWorkbook(excelFile);
		XSSFSheet sheetTestData = getExampleSheet(worksheetName, excelFile, workbook);

		return getExampleListFromWorksheet(excelFile, sheetTestData, headerRow, parameterNameColumn, headerMatcher,
				headerUnmatcher, columnType);
	}

	private static List<Map<String, String>> getExampleListFromWorksheet(FileInputStream excelFile,
			XSSFSheet sheetTestData, int headerRow, char parameterNameColumn, String headerMatcher,
			String headerUnmatcher, String columnType) throws IOException {
		// poi get row from 0, so 1st headerRow is at 0
		// by default, actualHeaderRow is below
		int actualHeaderRow = headerRow - 1;
		int actualParameterStartRow = headerRow;
		int columnStep = 1;
		if (TESTRESULT.equals(columnType)) {
			// because of input/expected/testresult row, the below -2
			actualParameterStartRow = headerRow + 1;
			columnStep = 3;
		} else if (EXPECTED.equals(columnType)) {
			actualParameterStartRow = headerRow + 1;
			columnStep = 2;
		}
		ArrayList<Map<String, String>> listTestSet = new ArrayList<>();
		// poi get column from 0, so Column A's Num is 0, 65 is A's ASCII code
		int parameterNameColumnNum = (int) parameterNameColumn - 65;

		XSSFRow rowHeader = sheetTestData.getRow(actualHeaderRow);
		HashMap<Integer, Integer> mapTestSetHeader = getHeaderMap(headerMatcher, headerUnmatcher, listTestSet,
				parameterNameColumnNum, rowHeader, columnStep);

		// Get ParameterNames HashMap
		HashMap<Integer, String> mapParameterName = getParameterNameMap(actualParameterStartRow, parameterNameColumnNum,
				sheetTestData);

		for (Map.Entry<Integer, String> aParameterName : mapParameterName.entrySet()) {
			int iRow = aParameterName.getKey();
			String strParameterName = aParameterName.getValue();
			XSSFRow rowCurrent = sheetTestData.getRow(iRow);

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
		sheetTestData.getWorkbook().close();
		excelFile.close();
		return listTestSet;
	}

	/**
	 * @param strParameterName
	 * @param rowCurrent
	 * @param mapTestSet
	 * @param iCol
	 */
	private static void putParameter(String strParameterName, XSSFRow rowCurrent, Map<String, String> mapTestSet,
			int iCol) {
		XSSFCell cellCurrent = rowCurrent.getCell(iCol);
		if (cellCurrent == null) {
			mapTestSet.put(strParameterName, "");
		} else {
			if (cellCurrent.getCellType().equals(CellType.STRING)) {
				mapTestSet.put(strParameterName, cellCurrent.getStringCellValue());
			} else if (cellCurrent.getCellType().equals(CellType.BLANK)) {
				mapTestSet.put(strParameterName, "");
			} else if (cellCurrent.getCellType().equals(CellType.NUMERIC)) {
				mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getNumericCellValue()));
			} else if (cellCurrent.getCellType().equals(CellType.BOOLEAN)) {
				mapTestSet.put(strParameterName, String.valueOf(cellCurrent.getBooleanCellValue()));
			} else {
				mapTestSet.put(strParameterName, cellCurrent.getRawValue());
			}
		}
	}
}
