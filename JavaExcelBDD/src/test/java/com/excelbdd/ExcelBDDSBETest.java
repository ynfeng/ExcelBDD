package com.excelbdd;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelBDDSBETest {

	static Stream<Map<String, String>> provideExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		return Behavior.getExampleStream(filePath, "SpecificationByExample", 1, 'F');
	}

	@ParameterizedTest(name = "Test{index}:{0}")
	@MethodSource("provideExampleList")
	void testParameterizedTestFromgetExampleStream(Map<String, String> parameterMap) throws IOException {
		assertNotNull(parameterMap);
		System.out.println("Header " + parameterMap.get("Header"));
		System.out.println("SheetName " + parameterMap.get("SheetName"));
		System.out.println("HeaderRow " + parameterMap.get("HeaderRow"));
		System.out.println("ParameterNameColumn " + parameterMap.get("ParameterNameColumn"));
		assertEquals("Scenario1", parameterMap.get("Header1Name"));
		assertEquals("V1.2", parameterMap.get("ParamName1InSet2Value"));
		assertEquals("", parameterMap.get("ParamName3Value"));

		assertEquals("3.0", parameterMap.get("MaxBlankThreshold"));
		System.out.println("HeaderMatcher " + parameterMap.get("HeaderMatcher"));
		assertEquals(true, parameterMap.get("Header").matches("Scenario.*"));
		assertEquals(false, parameterMap.get("Header").matches("V0.*"));

		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		int nHeaderRow = Double.valueOf(parameterMap.get("HeaderRow")).intValue();
		char charParameterNameColumn = parameterMap.get("ParameterNameColumn").charAt(0);
		System.out.println("ParameterNameColumn " + charParameterNameColumn);

		List<Map<String, String>> list;
		if (parameterMap.get("TestResultSwitch").equals("On")) {
			list = Behavior.getExampleListWithTestResult(filepath, parameterMap.get("SheetName"), nHeaderRow,
					charParameterNameColumn, parameterMap.get("HeaderMatcher"));
		} else if (parameterMap.get("ExpectedSwitch").equals("On")) {
			list = Behavior.getExampleListWithExpected(filepath, parameterMap.get("SheetName"), nHeaderRow,
					charParameterNameColumn, parameterMap.get("HeaderMatcher"));
		} else {
			list = Behavior.getExampleList(filepath, parameterMap.get("SheetName"), nHeaderRow, charParameterNameColumn,
					parameterMap.get("HeaderMatcher"), parameterMap.get("HeaderUnmatcher"));
		}

		System.out.println("ParameterCount " + parameterMap.get("ParameterCount"));
		assertEquals(TestWizard.getInt(parameterMap.get("ParameterCount")), list.get(0).size());
		System.out.println(list.get(0).toString());
		System.out.println(list.get(1).toString());
		System.out.println(list.get(2).toString());
		System.out.println(list.get(3).toString());

		// int testDataSetCount =
		// Double.valueOf(parameterMap.get("TestDataSetCount")).intValue();
		int testDataSetCount = TestWizard.getInt(parameterMap.get("TestDataSetCount"));
		assertEquals(testDataSetCount, list.size());

		assertEquals(parameterMap.get("FirstGridValue"), list.get(0).get("ParamName1"));
		assertEquals(parameterMap.get("ParamName1InSet2Value"), list.get(1).get("ParamName1"));
		assertEquals("V1.3", list.get(2).get("ParamName1"));
		assertEquals("V1.4", list.get(3).get("ParamName1"));

		assertEquals("V2.1", list.get(0).get("ParamName2"));
		assertEquals(parameterMap.get("ParamName2InSet2Value"), list.get(1).get("ParamName2"));

		assertEquals("", list.get(0).get("ParamName3"));
		assertEquals("", list.get(1).get("ParamName3"));
		assertEquals("", list.get(2).get("ParamName3"));
		assertEquals("", list.get(3).get("ParamName3"));

		assertEquals("2021/4/30", list.get(0).get("ParamName4"));
		assertEquals("false", list.get(1).get("ParamName4"));
		assertEquals("true", list.get(2).get("ParamName4"));
		assertEquals(parameterMap.get("LastGridValue"), list.get(3).get("ParamName4"));
	}

	@Test
	void testgetInt() {
		assertEquals(5, TestWizard.getInt("5.5666"));
		assertEquals(5, TestWizard.getInt("5"));
		assertEquals(5, TestWizard.getInt("5.99999"));
	}

	@Test
	void testBDDExcelPath2() throws IOException {
		String ExcelFilePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		File f = new File(ExcelFilePath);
		assertTrue(f.exists());
	}
}
