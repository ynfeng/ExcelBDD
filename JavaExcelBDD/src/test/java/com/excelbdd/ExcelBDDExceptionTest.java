package com.excelbdd;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

class ExcelBDDExceptionTest {

    static Stream<Map<String, String>> provideWrongFileExampleList() throws IOException {
        String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        return Behavior.getExamples(filePath, "WrongFile", "Wrong").stream();
    }

    @ParameterizedTest(name = "#{index}-TestException: {0}")
    @MethodSource("provideWrongFileExampleList")
    void testGetWrongFileExampleList(Map<String, String> mapParams) {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
            + mapParams.get("ExcelFileName");
        int headerRow = TestWizard.getInt(mapParams.get("HeaderRow"));
        char parameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
        Throwable exception = assertThrows(IOException.class, () -> {
            List<Map<String, String>> targetlist = Behavior.getExampleListWithExpected(filepath,
                mapParams.get("SheetName"), headerRow);
        });

        System.out.println(exception.toString());
        System.out.println(exception.getClass().getSimpleName());
        assertEquals(mapParams.get("ExcelFileNameExpected"), exception.getClass().getSimpleName());
        assertTrue(exception.getMessage().contains(mapParams.get("SheetNameExpected")));
    }

    static Stream<Map<String, String>> provideWrongOtherExampleList() throws IOException {
        String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
        List<Map<String, String>> list = Behavior.getExampleListWithExpected(filePath, "WrongOther", 1);
        return list.stream();
    }

    @ParameterizedTest(name = "#{index}-TestException: {0}")
    @MethodSource("provideWrongOtherExampleList")
    void testGetWrongOtherExampleList(Map<String, String> mapParams) throws IOException {
        String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/"
            + mapParams.get("ExcelFileName");
        int headerRow = TestWizard.getInt(mapParams.get("HeaderRow"));
        char parameterNameColumn = mapParams.get("ParameterNameColumn").charAt(0);
        List<Map<String, String>> targetList = Behavior.getExamples(filepath,
            mapParams.get("SheetName"), headerRow);

        assertEquals(TestWizard.getInt(mapParams.get("HeaderRowExpected")), targetList.size());
        assertEquals(TestWizard.getInt(mapParams.get("ParameterNameColumnExpected")), targetList.get(0).size());
    }
}
