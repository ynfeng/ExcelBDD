package com.excelbdd.cucumber;

import static org.junit.jupiter.api.Assertions.*;
import java.util.List;
import java.util.Map;

import io.cucumber.java.en.Given;
import io.cucumber.java.en.When;

public class BDDSteps {

	@Given("the following animals:")
	public void the_following_animals(List<String> listAnimal) {
		assertEquals("cow", listAnimal.get(0));
		for (String animal : listAnimal) {
			// prints the elements of the List
			System.out.println(animal);
		}
	}

	@Given("the cities basic status:")
	public void the_cities_basic_status(List<List<String>> tableCity) {
		for (List<String> city : tableCity) {
			for (String cities : city) {
				// prints the elements of the List
				System.out.println(cities);
			}
		}
	}

	@When("key and value list:")
	public void key_and_value_list(io.cucumber.datatable.DataTable dataTable) {
		assertEquals("header1", dataTable.cell(0, 0));
		List<Map<String, String>> mapList = dataTable.asMaps();
		for (Map<String, String> map : mapList) {
			System.out.println("===========");
			for (Map.Entry<String, String> mapEntry : map.entrySet()) {
				System.out.print(mapEntry.getKey());
				System.out.println(mapEntry.getValue());
			}
			System.out.println("map.keySet()");
			for (String header : map.keySet()) {
				System.out.print(header);
				System.out.println(map.get(header));
			}
		}

		System.out.println("=======dataTable.asLists()=======");
		List<List<String>> listList = dataTable.asLists();
		for (List<String> list : listList) {
			System.out.println("===========");
			for (String string : list) {
				// prints the elements of the List
				System.out.println(string);
			}
		}
	}

}
