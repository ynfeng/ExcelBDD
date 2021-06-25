package com.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

public class ExcelBDDSBETest {

    static Stream<Map<String, String>> provideExampleList() throws IOException {
        String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        return Behavior.getExampleStream(filePath, "SpecificationByExample", 1);
    }

    @ParameterizedTest(name = "Test{index}:{0}")
    @MethodSource("provideExampleList")
    void testParameterizedTestFromgetExampleStream(Map<String, String> mapParams) throws IOException {
        assertNotNull(mapParams);
        System.out.println("Header " + mapParams.get("Header"));
        System.out.println("SheetName " + mapParams.get("SheetName"));
        System.out.println("HeaderRow " + mapParams.get("HeaderRow"));
        System.out.println("ParameterNameColumn " + mapParams.get("ParameterNameColumn"));
        assertEquals("Scenario1", mapParams.get("Header1Name"));
        assertEquals("V1.2", mapParams.get("ParamName1InSet2Value"));
        assertEquals("", mapParams.get("ParamName3Value"));

        assertEquals("3.0", mapParams.get("MaxBlankThreshold"));
        System.out.println("HeaderMatcher " + mapParams.get("HeaderMatcher"));
        assertEquals(true, mapParams.get("Header").matches("Scenario.*"));
        assertEquals(false, mapParams.get("Header").matches("V0.*"));


        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        int nHeaderRow = Double.valueOf(mapParams.get("HeaderRow")).intValue();
        char charParameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
        System.out.println("ParameterNameColumn " + charParameterNameColumn);

        List<Map<String, String>> list;
        if (mapParams.get("TestResultSwitch").equals("On")) {
            list = Behavior.getExampleListWithTestResult(filepath, mapParams.get("SheetName"), nHeaderRow,
                mapParams.get("HeaderMatcher"));
        } else if (mapParams.get("ExpectedSwitch").equals("On")) {
            list = Behavior.getExampleListWithExpected(filepath, mapParams.get("SheetName"), nHeaderRow,
                mapParams.get("HeaderMatcher"));
        } else {
            list = Behavior.getExamples(filepath, mapParams.get("SheetName"), nHeaderRow,
                mapParams.get("HeaderMatcher"), mapParams.get("HeaderUnmatcher"));
        }

        System.out.println("ParameterCount " + mapParams.get("ParameterCount"));
        assertEquals(TestWizard.getInt(mapParams.get("ParameterCount")), list.get(0).size());
        System.out.println(list.get(0).toString());
        System.out.println(list.get(1).toString());
        System.out.println(list.get(2).toString());
        System.out.println(list.get(3).toString());

        // int testDataSetCount =
        // Double.valueOf(mapParams.get("TestDataSetCount")).intValue();
        int testDataSetCount = TestWizard.getInt(mapParams.get("TestDataSetCount"));
        assertEquals(testDataSetCount, list.size());

        assertEquals(mapParams.get("FirstGridValue"), list.get(0).get("ParamName1"));
        assertEquals(mapParams.get("ParamName1InSet2Value"), list.get(1).get("ParamName1"));
        assertEquals("V1.3", list.get(2).get("ParamName1"));
        assertEquals("V1.4", list.get(3).get("ParamName1"));

        assertEquals("V2.1", list.get(0).get("ParamName2"));
        assertEquals(mapParams.get("ParamName2InSet2Value"), list.get(1).get("ParamName2"));

        assertEquals("", list.get(0).get("ParamName3"));
        assertEquals("", list.get(1).get("ParamName3"));
        assertEquals("", list.get(2).get("ParamName3"));
        assertEquals("", list.get(3).get("ParamName3"));

        assertEquals("2021/4/30", list.get(0).get("ParamName4"));
        assertEquals("false", list.get(1).get("ParamName4"));
        assertEquals("true", list.get(2).get("ParamName4"));
        assertEquals(mapParams.get("LastGridValue"), list.get(3).get("ParamName4"));
    }

    @Test
    void testgetInt() {
        assertEquals(5, TestWizard.getInt("5.5666"));
        assertEquals(5, TestWizard.getInt("5"));
        assertEquals(5, TestWizard.getInt("5.99999"));
    }

    @Test
    void testBDDExcelPath2() {
        String ExcelFilePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        File f = new File(ExcelFilePath);
        assertTrue(f.exists());
    }
}
