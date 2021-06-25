package com.excelbdd;

import static org.hamcrest.CoreMatchers.instanceOf;
import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.jupiter.api.Assertions.fail;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

@SuppressWarnings("NonAsciiCharacters")
class BehaviorTest {

    @Test
    void sheet不存在时应该抛异常() {
        String filePath = getExcelPath("NoSheet.xlsx");
        try {
            Behavior.getExamples(filePath, "not-exists");
            fail();
        } catch (Exception e) {
            assertThat(e, instanceOf(IOException.class));
        }
    }

    @Test
    void 可以读取有空行的excel() throws IOException {
        String filePath = getExcelPath("HasEmptyRow.xlsx");

        List<Map<String, String>> result = Behavior.getExamples(filePath, "DataTableBDD");

        assertThat(result.size(), is(3));
    }

    private String getExcelPath(String excelName) {
        return TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/" + excelName;
    }
}
