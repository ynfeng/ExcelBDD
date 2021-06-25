package com.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.util.Collection;
import java.util.Map;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

@RunWith(Parameterized.class)
public class ExcelBDDParameterizedTest {

    Map<String, String> mapParams;

    @Parameters
    public static Collection<Object[]> prepareData() throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        return Behavior.getExampleCollection(filepath, "SpecificationByExample", 1);
    }

    public ExcelBDDParameterizedTest(Map<String, String> map) {
        this.mapParams = map;
    }

    @Test
    public void testZMExcelParameterizedTest() {

        System.out.println("Header " + mapParams.get("Header"));
        System.out.println("SheetName " + mapParams.get("SheetName"));
        System.out.println("HeaderRow " + mapParams.get("HeaderRow"));

        System.out.println(mapParams.get("MaxBlankThreshold"));
        System.out.println(mapParams.get("HeaderMatcher"));
        System.out.println(mapParams.get("ParameterCount"));
        System.out.println("ParameterNameColumn " + mapParams.get("ParameterNameColumn"));

        assertEquals("Scenario1", mapParams.get("Header1Name"));
        assertEquals("V1.2", mapParams.get("ParamName1InSet2Value"));
        assertEquals("V2.2", mapParams.get("ParamName2InSet2Value"));
        assertEquals("", mapParams.get("ParamName3Value"));
    }
}
