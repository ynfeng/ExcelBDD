package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class ExcelBDDExceptionTest {

	static Stream<Map<String, String>> provideWrongFileExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		return Behavior.getExampleStream(filePath, "WrongFile", "Wrong");
	}

	@ParameterizedTest(name = "#{index}-TestException: {0}")
	@MethodSource("provideWrongFileExampleList")
	void testGetWrongFileExampleList(Map<String, String> parameterMap) throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
				+ parameterMap.get("ExcelFileName");
		int headerRow = TestWizard.getInt(parameterMap.get("HeaderRow"));
		char parameterNameColumn = parameterMap.get("ParameterNameColumn").charAt(0);
		Throwable exception = assertThrows(IOException.class, () -> {
			List<Map<String, String>> targetList = Behavior.getExampleList(filepath,parameterMap.get("SheetName"));
		});
		
		System.out.println(exception.toString());
		System.out.println(exception.getClass().getSimpleName());
		assertEquals(parameterMap.get("ExcelFileNameExpected"), exception.getClass().getSimpleName());
		System.out.println(exception.getMessage());
		assertTrue(exception.getMessage().contains(parameterMap.get("SheetNameExpected")));
	}
	
	static Stream<Map<String, String>> provideWrongOtherExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithExpected(filePath, "WrongOther", 1, 'C');
		return list.stream();
	}

	@ParameterizedTest(name = "#{index}-TestException: {0}")
	@MethodSource("provideWrongOtherExampleList")
	void testGetWrongOtherExampleList(Map<String, String> parameterMap) throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
				+ parameterMap.get("ExcelFileName");
		int headerRow = TestWizard.getInt(parameterMap.get("HeaderRow"));
		char parameterNameColumn = parameterMap.get("ParameterNameColumn").charAt(0);
		List<Map<String, String>> targetList = Behavior.getExampleList(filepath,
					parameterMap.get("SheetName"), headerRow, parameterNameColumn);
		
		assertEquals(TestWizard.getInt(parameterMap.get("HeaderRowExpected")),targetList.size());
		assertEquals(TestWizard.getInt(parameterMap.get("ParameterNameColumnExpected")),targetList.get(0).size());
	}
	
	
}
