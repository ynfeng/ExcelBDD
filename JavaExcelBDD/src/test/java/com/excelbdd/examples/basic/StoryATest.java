package com.excelbdd.examples.basic;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.Map;
import java.util.stream.Stream;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import com.excelbdd.Behavior;
import com.excelbdd.TestWizard;

class StoryATest {

	static Stream<Map<String, String>> provideExampleList() throws IOException {
		String filePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
		return Behavior.getExampleStream(filePath, "StoryExample1");
	}

	@ParameterizedTest(name = "Test{index}:{0}")
	@MethodSource("provideExampleList")
	void testgetCellValue(Map<String, String> parameterMap) throws IOException {
		assertNotNull(parameterMap.get("Header"));
		TestWizard.showMap(parameterMap);
	}
}
