package com.excelbdd;

import java.io.File;
import java.io.IOException;
import java.util.Map;

public class TestWizard {
	public static final String ANY_MATCHER = ".*";
	public static final String NEVER_MATCHED_STRING = "i_m_p_o_s_i_b_l_e_matcher";

	private TestWizard() {
	}

	public static String getExcelBDDStartPath(String childPath) throws IOException {
		String absolutePath = new File(".").getAbsolutePath();
		if(absolutePath.lastIndexOf(childPath) >= 0) {
			return absolutePath.substring(0, absolutePath.lastIndexOf(childPath));
		}else {
			throw new IOException(childPath + " is not in " + absolutePath);
		}
	}

	public static void showMap(Map<String, String> parameterMap) {
		System.out.println(String.format("=======Header: %s=====", parameterMap.get("Header")));
		for (Map.Entry<String, String> param : parameterMap.entrySet()) {
			System.out.println(String.format("%s --- %s", param.getKey(), param.getValue()));
		}
	}

	public static int getInt(String string) {
		return Double.valueOf(string).intValue();
	}

	public static String makeMatcherString(String matcher) {
		return ANY_MATCHER + matcher + ANY_MATCHER;
	}
}
