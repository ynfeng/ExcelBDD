/**
 *
 */

package com.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;


class ExcelBDDSBTTest {
    static Stream<Map<String, String>> provideExampleListWithExpected() throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        return Behavior.getExamples(filepath, "Expected1", TestWizard.ANY_MATCHER, TestWizard.NEVER_MATCHED_STRING).stream();
    }

    /**
     * Test method for {@link #getExampleListWithExpected(String, String, int)}.
     */
    @ParameterizedTest(name = "#{index}-TestExpected : {0}")
    @MethodSource("provideExampleListWithExpected")
    void testGetExampleListWithExpected(Map<String, String> mapParams) {
        assertNotNull(mapParams);
        TestWizard.showMap(mapParams);
        assertTrue(mapParams.get("ParamName1").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3").isEmpty());
        assertEquals(false, mapParams.get("ParamName4").isEmpty());

        assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3Expected").isEmpty());
        assertEquals(false, mapParams.get("ParamName4Expected").isEmpty());

        assertNull(mapParams.get("ParamName1TestResult"));
        assertNull(mapParams.get("ParamName2TestResult"));
        assertNull(mapParams.get("ParamName3TestResult"));
        assertNull(mapParams.get("ParamName4TestResult"));
    }

    static Stream<Map<String, String>> provideExampleListWithExpectedByMatcher() throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        List<Map<String, String>> list = Behavior.getExampleListWithExpected(filepath, "Expected1", 1, "Scenario");
        return list.stream();
    }

    /**
     * Test method for {@link #getExampleListWithExpected(String, String, int, String)}.
     */
    @ParameterizedTest(name = "#{index}-TestExpected : {0}")
    @MethodSource("provideExampleListWithExpectedByMatcher")
    void testGetExampleListWithExpectedByMatcher(Map<String, String> mapParams) {
        assertNotNull(mapParams);
        TestWizard.showMap(mapParams);
        assertTrue(mapParams.get("ParamName1").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3").isEmpty());
        assertEquals(false, mapParams.get("ParamName4").isEmpty());

        assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3Expected").isEmpty());
        assertEquals(false, mapParams.get("ParamName4Expected").isEmpty());

        assertNull(mapParams.get("ParamName1TestResult"));
        assertNull(mapParams.get("ParamName2TestResult"));
        assertNull(mapParams.get("ParamName3TestResult"));
        assertNull(mapParams.get("ParamName4TestResult"));
    }

    static Stream<Map<String, String>> provideExampleListWithTestResultStringStringIntChar() throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        List<Map<String, String>> list = Behavior.getExampleListWithTestResult(filepath, "SBTSheet1", 1);
        return list.stream();
    }

    /**
     * Test method for {@link #getExampleListWithTestResult(String, String, int)}.
     */
    @ParameterizedTest(name = "TestResult{index}:{0}")
    @MethodSource("provideExampleListWithTestResultStringStringIntChar")
    void testGetExampleListWithTestResultStringStringIntChar(Map<String, String> mapParams) {
        TestWizard.showMap(mapParams);
        assertTrue(mapParams.get("ParamName1").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3").isEmpty());
        assertEquals(false, mapParams.get("ParamName4").isEmpty());

        assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3Expected").isEmpty());
        assertEquals(false, mapParams.get("ParamName4Expected").isEmpty());

        assertNotNull(mapParams.get("ParamName1TestResult"));
        assertNotNull(mapParams.get("ParamName2TestResult"));
        assertNotNull(mapParams.get("ParamName3TestResult"));
        assertNotNull(mapParams.get("ParamName4TestResult"));

        assertEquals("pass", mapParams.get("ParamName1TestResult"));
        assertEquals("pass", mapParams.get("ParamName2TestResult"));
        assertEquals("pass", mapParams.get("ParamName3TestResult"));
        assertEquals("pass", mapParams.get("ParamName4TestResult"));
    }

    static Stream<Map<String, String>> provideExampleListWithTestResultStringStringIntCharString() throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        List<Map<String, String>> list = Behavior.getExampleListWithTestResult(filepath, "SBTSheet1", 1, "Scenario");
        return list.stream();
    }

    /**
     * Test method for {@link #getExampleListWithTestResult(String, String, int, String)}.
     */
    @ParameterizedTest(name = "TestResult{index}:{0}")
    @MethodSource("provideExampleListWithTestResultStringStringIntCharString")
    void testGetExampleListWithTestResultStringStringIntCharString(Map<String, String> mapParams) {
        TestWizard.showMap(mapParams);
        assertTrue(mapParams.get("ParamName1").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3").isEmpty());
        assertEquals(false, mapParams.get("ParamName4").isEmpty());

        assertTrue(mapParams.get("ParamName1Expected").startsWith("V1."));
        assertTrue(mapParams.get("ParamName2Expected").startsWith("V2."));
        assertEquals(true, mapParams.get("ParamName3Expected").isEmpty());
        assertEquals(false, mapParams.get("ParamName4Expected").isEmpty());

        assertNotNull(mapParams.get("ParamName1TestResult"));
        assertNotNull(mapParams.get("ParamName2TestResult"));
        assertNotNull(mapParams.get("ParamName3TestResult"));
        assertNotNull(mapParams.get("ParamName4TestResult"));

        assertEquals("pass", mapParams.get("ParamName1TestResult"));
        assertEquals("pass", mapParams.get("ParamName2TestResult"));
        assertEquals("pass", mapParams.get("ParamName3TestResult"));
        assertEquals("pass", mapParams.get("ParamName4TestResult"));
    }
}
