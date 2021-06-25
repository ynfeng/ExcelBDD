package com.excelbdd;

import static org.junit.Assert.*;
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
	
	Map<String, String> parameterMap;
	
	@Parameters
    public static Collection<Object[]> prepareData() throws IOException
    {
    	String filepath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + "BDDExcel/ExcelBDD.xlsx";
    	return Behavior.getExampleCollection(filepath, "SpecificationByExample", 1, 'F');
    }
	
	public ExcelBDDParameterizedTest(Map<String, String> map){
		this.parameterMap = map;
	}

	@Test
	public void testZMExcelParameterizedTest() {
		
		System.out.println("Header " + parameterMap.get("Header"));
		System.out.println("SheetName " + parameterMap.get("SheetName"));
		System.out.println("HeaderRow " + parameterMap.get("HeaderRow"));

		System.out.println(parameterMap.get("MaxBlankThreshold"));
		System.out.println(parameterMap.get("HeaderMatcher"));
		System.out.println(parameterMap.get("ParameterCount"));		
		System.out.println("ParameterNameColumn " + parameterMap.get("ParameterNameColumn"));
		
		assertEquals("Scenario1", parameterMap.get("Header1Name"));
		assertEquals("V1.2", parameterMap.get("ParamName1InSet2Value"));
		assertEquals("V2.2", parameterMap.get("ParamName2InSet2Value"));
		assertEquals("", parameterMap.get("ParamName3Value"));
	}
}
