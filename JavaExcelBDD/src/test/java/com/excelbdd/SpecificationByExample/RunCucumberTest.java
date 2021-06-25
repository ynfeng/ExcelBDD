package com.excelbdd.SpecificationByExample;

import io.cucumber.junit.CucumberOptions;
import io.cucumber.junit.Cucumber;
import org.junit.runner.RunWith;

@RunWith(Cucumber.class)
// @CucumberOptions(plugin = {"pretty", "html:target/cucumber-report.html"})
@CucumberOptions(plugin = { "pretty" })
public class RunCucumberTest {

}
