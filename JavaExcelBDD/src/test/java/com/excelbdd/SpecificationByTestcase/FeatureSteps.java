package com.excelbdd.SpecificationByTestcase;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import com.excelbdd.TestWizard;
import com.excelbdd.Behavior;

import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import io.cucumber.java.en.When;

public class FeatureSteps {
	String excelFilePath;
	String sheetName;
	int headerRow;
	char parameterNameColumn;
	String headerMatcher;
	List<Map<String, String>> list;

	@Given("The Excel file is {string}")
	public void the_excel_file_is(String string) throws IOException {
		excelFilePath = TestWizard.getExcelBDDStartPath("JavaExcelBDD") + string;
	}

	@Given("The Sheet name is {string}")
	public void the_sheet_name_is(String string) {
		sheetName = string;
	}

	@Given("Header Row is {int}")
	public void header_row_is(int int1) {
		headerRow = int1;
	}

	@Given("Parameter Column is {string}")
	public void parameter_column_is(String string) {
		parameterNameColumn = string.charAt(0);
	}

	@Given("Matcher is {string}")
	public void matcher_is(String string) {
		headerMatcher = string;
	}

	@When("invoke get test data from excel")
	public void invoke_get_test_data_from_excel() throws IOException {
		list = Behavior.getExampleListWithTestResult(excelFilePath, sheetName, headerRow, parameterNameColumn);
		for (Map<String, String> map : list) {
			System.out.println(map.get("Header") + " ========== ");
			for (Map.Entry<String, String> mapEntry : map.entrySet()) {
				System.out.print(mapEntry.getKey() + " --- ");
				System.out.println(mapEntry.getValue());
			}
		}
	}

	@When("invoke get test data from excel according to Matcher")
	public void invoke_get_test_data_from_excel_according_to_matcher() throws IOException {
		list = Behavior.getExampleListWithTestResult(excelFilePath, sheetName, headerRow, parameterNameColumn,
				headerMatcher);
		for (Map<String, String> map : list) {
			System.out.println(map.get("Header") + " ========== ");
			for (Map.Entry<String, String> mapEntry : map.entrySet()) {
				System.out.print(mapEntry.getKey() + " --- ");
				System.out.println(mapEntry.getValue());
			}
		}
	}

	@Then("a testset list is got, which count is {int}")
	public void a_testset_list_is_got_which_count_is(int int1) {
		assertEquals(int1, list.size());
	}

	@Then("The Header of No. {int} set is {string}")
	public void the_header_of_1st_set_is(int int1, String string) {
		assertEquals(string, list.get(int1 - 1).get("Header"));
	}

	@Then("Input value of Variable {string} of No. {int} set is {string}")
	public void input_value_of_variable_is(String string, int int1, String string2) {
		assertEquals(string2, list.get(int1 - 1).get(string));
	}

	@Then("Expected value of Variable {string} of No. {int} set is {string}")
	public void expected_value_of_variable_is(String string, int int1, String string2) {
		assertEquals(string2, list.get(int1 - 1).get(string + "Expected"));
	}

	@Then("Test Result value of Variable {string} of No. {int} set is {string}")
	public void test_result_value_of_variable_is(String string, int int1, String string2) {
		assertEquals(string2, list.get(int1 - 1).get(string + "TestResult"));
	}

	@Then("The 1st data table is:")
	public void the_data_table_is(io.cucumber.datatable.DataTable dataTable) {
		List<Map<String, String>> mapList = dataTable.asMaps();
		int i = 0;
		for (Map<String, String> map : mapList) {
			System.out.println("===========");
			for (Map.Entry<String, String> mapEntry : map.entrySet()) {
				System.out.print(mapEntry.getKey() + " --- ");
				System.out.println(mapEntry.getValue());
			}
			if (map.get("Input") == null) {
				assertEquals("", list.get(i).get(map.get("ParameterName")));
			} else {
				assertEquals(map.get("Input"), list.get(i).get(map.get("ParameterName")));
			}

			if (map.get("Expected") == null) {
				assertEquals("", list.get(i).get(map.get("ParameterName") + "Expected"));
			} else {
				assertEquals(map.get("Expected"), list.get(i).get(map.get("ParameterName") + "Expected"));
			}
			assertEquals("pass", map.get("TestResult"));
			assertEquals(map.get("TestResult"), list.get(i).get(map.get("ParameterName") + "TestResult"));
		}
	}
}
