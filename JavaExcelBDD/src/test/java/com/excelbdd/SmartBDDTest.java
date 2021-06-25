package com.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class SmartBDDTest {

    static Stream<Map<String, String>> provideExampleList() throws IOException {
        String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        return Behavior.getExamples(filePath, "SmartBDD").stream();
    }

    @ParameterizedTest(name = "Test{index}:{0}")
    @MethodSource("provideExampleList")
    void testgetSmartExampleStream(Map<String, String> mapParams) throws IOException {
        String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
            + mapParams.get("ExcelFileName");
        System.out.println("Header " + mapParams.get("Header"));
        System.out.println("SheetName " + mapParams.get("SheetName"));

        List<Map<String, String>> list = Behavior.getExamples(filePath, mapParams.get("SheetName"),
            mapParams.get("HeaderMatcher"), mapParams.get("HeaderUnmatcher"));
        assertNotNull(list);

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

}
