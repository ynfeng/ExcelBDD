package com.excelbdd;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

public class Behavior {
    private static final String SIMPLE = "SIMPLE";
    private static final String TESTRESULT = "TESTRESULT";
    private static final String EXPECTED = "EXPECTED";

    private Behavior() {
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName) throws IOException {
        return getExamples(excelPath, worksheetName, TestWizard.ANY_MATCHER);
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName, String headerMatcher) throws IOException {
        return getExamples(excelPath, worksheetName, headerMatcher, TestWizard.NEVER_MATCHED_STRING);
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName, String headerMatcher, String headerUnmatcher) throws IOException {
        Excel excel = Excel.open(excelPath);
        Sheet sheet = excel.openSheet(worksheetName);

        try {
            return sheet.getExamples(headerMatcher, headerUnmatcher);
        } finally {
            sheet.close();
            excel.close();
        }
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName, int headerRow) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, TestWizard.ANY_MATCHER,
            TestWizard.NEVER_MATCHED_STRING, SIMPLE);
    }

    public static Stream<Map<String, String>> getExampleStream(String excelPath, String worksheetName, int headerRow) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, TestWizard.ANY_MATCHER,
            TestWizard.NEVER_MATCHED_STRING, SIMPLE).stream();
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName, int headerRow, String headerMatcher, String headerUnmatcher) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, headerMatcher, headerUnmatcher,
            SIMPLE);
    }

    public static Collection<Object[]> getExampleCollection(String excelPath, String worksheetName, int headerRow) throws IOException {
        Collection<Object[]> collectionTestData = new ArrayList<>();
        List<Map<String, String>> listTestData = getExamples(excelPath, worksheetName, headerRow,
            TestWizard.ANY_MATCHER, TestWizard.NEVER_MATCHED_STRING, SIMPLE);
        for (Map<String, String> map : listTestData) {
            Object[] arrayObj = {map};
            collectionTestData.add(arrayObj);
        }
        return collectionTestData;
    }

    public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName, int headerRow) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, TestWizard.ANY_MATCHER,
            TestWizard.NEVER_MATCHED_STRING, EXPECTED);
    }

    public static List<Map<String, String>> getExampleListWithExpected(String excelPath, String worksheetName, int headerRow, String headerMatcher) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, headerMatcher,
            TestWizard.NEVER_MATCHED_STRING, EXPECTED);
    }

    public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName, int headerRow) throws IOException {
        return getExampleListWithTestResult(excelPath, worksheetName, headerRow, TestWizard.ANY_MATCHER);
    }

    public static List<Map<String, String>> getExampleListWithTestResult(String excelPath, String worksheetName, int headerRow, String headerMatcher) throws IOException {
        return getExamples(excelPath, worksheetName, headerRow, headerMatcher,
            TestWizard.NEVER_MATCHED_STRING, TESTRESULT);
    }

    public static List<Map<String, String>> getExamples(String excelPath, String worksheetName, int headerRow, String headerMatcher, String headerUnmatcher, String columnType) throws IOException {
        Excel excel = Excel.open(excelPath);
        Sheet sheet = excel.openSheet(worksheetName);

        try {
            return sheet.getExamples(headerRow, headerMatcher, headerUnmatcher);
        } finally {
            sheet.close();
            excel.close();
        }
    }

}
