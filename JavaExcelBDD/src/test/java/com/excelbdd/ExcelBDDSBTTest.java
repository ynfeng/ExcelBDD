/**
 * 
 */
package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class ExcelBDDSBTTest {
	static Stream<Map<String, String>> provideExampleListWithExpected() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		return Behavior.getExampleStream(filepath, "Expected1", TestWizard.ANY_MATCHER,
				TestWizard.NEVER_MATCHED_STRING);
	}

	/**
	 * Test method for
	 * {@link com.excelbdd.Behavior#getExampleListWithExpected(java.lang.String, java.lang.String, int, char)}.
	 */
	@ParameterizedTest(name = "#{index}-TestExpected : {0}")
	@MethodSource("provideExampleListWithExpected")
	void testGetExampleListWithExpected(Map<String, String> parameterMap) {
		assertNotNull(parameterMap);
		TestWizard.showMap(parameterMap);
		assertTrue(parameterMap.get("ParamName1").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2").startsWith("V2."));
		assertEquals(true, parameterMap.get("ParamName3").isEmpty());
		assertEquals(false, parameterMap.get("ParamName4").isEmpty());

		assertTrue(parameterMap.get("ParamName1Expected").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2Expected").startsWith("V2."));
		assertTrue(parameterMap.get("ParamName3Expected").isEmpty());
		assertFalse(parameterMap.get("ParamName4Expected").isEmpty());

		assertNull(parameterMap.get("ParamName1TestResult"));
		assertNull(parameterMap.get("ParamName2TestResult"));
		assertNull(parameterMap.get("ParamName3TestResult"));
		assertNull(parameterMap.get("ParamName4TestResult"));
	}

	static Stream<Map<String, String>> provideExampleListWithExpectedByMatcher() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithExpected(filepath, "Expected1", 1, 'B', "Scenario");
		return list.stream();
	}

	/**
	 * Test method for
	 * {@link com.excelbdd.Behavior#getExampleListWithExpected(java.lang.String, java.lang.String, int, char, java.lang.String)}.
	 */
	@ParameterizedTest(name = "#{index}-TestExpected : {0}")
	@MethodSource("provideExampleListWithExpectedByMatcher")
	void testGetExampleListWithExpectedByMatcher(Map<String, String> parameterMap) {
		assertNotNull(parameterMap);
		TestWizard.showMap(parameterMap);
		assertTrue(parameterMap.get("ParamName1").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2").startsWith("V2."));
		assertEquals(true, parameterMap.get("ParamName3").isEmpty());
		assertEquals(false, parameterMap.get("ParamName4").isEmpty());

		assertTrue(parameterMap.get("ParamName1Expected").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2Expected").startsWith("V2."));
		assertTrue(parameterMap.get("ParamName3Expected").isEmpty());
		assertFalse(parameterMap.get("ParamName4Expected").isEmpty());

		assertNull(parameterMap.get("ParamName1TestResult"));
		assertNull(parameterMap.get("ParamName2TestResult"));
		assertNull(parameterMap.get("ParamName3TestResult"));
		assertNull(parameterMap.get("ParamName4TestResult"));
	}

	static Stream<Map<String, String>> provideExampleListWithTestResultStringStringIntChar() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		List<Map<String, String>> list = Behavior.getExampleListWithTestResult(filepath, "SBTSheet1", 1, 'B');
		return list.stream();
	}

	/**
	 * Test method for
	 * {@link com.excelbdd.Behavior#getExampleListWithTestResult(java.lang.String, java.lang.String, int, char)}.
	 */
	@ParameterizedTest(name = "TestResult{index}:{0}")
	@MethodSource("provideExampleListWithTestResultStringStringIntChar")
	void testGetExampleListWithTestResultStringStringIntChar(Map<String, String> parameterMap) {
		TestWizard.showMap(parameterMap);
		assertTrue(parameterMap.get("ParamName1").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2").startsWith("V2."));
		assertEquals(true, parameterMap.get("ParamName3").isEmpty());
		assertEquals(false, parameterMap.get("ParamName4").isEmpty());

		assertTrue(parameterMap.get("ParamName1Expected").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2Expected").startsWith("V2."));
		assertTrue(parameterMap.get("ParamName3Expected").isEmpty());
		assertFalse(parameterMap.get("ParamName4Expected").isEmpty());

		assertNotNull(parameterMap.get("ParamName1TestResult"));
		assertNotNull(parameterMap.get("ParamName2TestResult"));
		assertNotNull(parameterMap.get("ParamName3TestResult"));
		assertNotNull(parameterMap.get("ParamName4TestResult"));

		assertEquals("pass", parameterMap.get("ParamName1TestResult"));
		assertEquals("pass", parameterMap.get("ParamName2TestResult"));
		assertEquals("pass", parameterMap.get("ParamName3TestResult"));
		assertEquals("pass", parameterMap.get("ParamName4TestResult"));
	}

	static Stream<Map<String, String>> provideExampleListWithTestResultStringStringIntCharString() throws IOException {
		String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		return Behavior.getExampleStream(filepath, "SBTSheet1", "Scenario");
	}

	/**
	 * Test method for
	 * {@link com.excelbdd.Behavior#getExampleListWithTestResult(java.lang.String, java.lang.String, int, char, java.lang.String)}.
	 */
	@ParameterizedTest(name = "TestResult{index}:{0}")
	@MethodSource("provideExampleListWithTestResultStringStringIntCharString")
	void testGetExampleListWithTestResultStringStringIntCharString(Map<String, String> parameterMap) {
		TestWizard.showMap(parameterMap);
		assertTrue(parameterMap.get("ParamName1").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2").startsWith("V2."));
		assertEquals(true, parameterMap.get("ParamName3").isEmpty());
		assertEquals(false, parameterMap.get("ParamName4").isEmpty());

		assertTrue(parameterMap.get("ParamName1Expected").startsWith("V1."));
		assertTrue(parameterMap.get("ParamName2Expected").startsWith("V2."));
		assertTrue(parameterMap.get("ParamName3Expected").isEmpty());
		assertFalse(parameterMap.get("ParamName4Expected").isEmpty());

		assertNotNull(parameterMap.get("ParamName1TestResult"));
		assertNotNull(parameterMap.get("ParamName2TestResult"));
		assertNotNull(parameterMap.get("ParamName3TestResult"));
		assertNotNull(parameterMap.get("ParamName4TestResult"));

		assertEquals("pass", parameterMap.get("ParamName1TestResult"));
		assertEquals("pass", parameterMap.get("ParamName2TestResult"));
		assertEquals("pass", parameterMap.get("ParamName3TestResult"));
		assertEquals("pass", parameterMap.get("ParamName4TestResult"));
	}
}
