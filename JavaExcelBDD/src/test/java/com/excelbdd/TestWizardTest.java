package com.excelbdd;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import org.junit.jupiter.api.Test;

class TestWizardTest {

	@Test
	void testgetExcelBDDStartPath() throws IOException {
		Throwable exception = assertThrows(IOException.class, () -> {
			@SuppressWarnings("unused")
			String path = TestWizard.getExcelBDDStartPath("NotExist");
		});

		assertTrue(exception.getMessage().contains("NotExist is not in"));
	}
}
