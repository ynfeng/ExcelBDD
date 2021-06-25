package com.excelbdd;

import java.io.File;
import java.util.Map;

public class TestWizard {
    static final String ANY_MATCHER = ".*";
    public static final String NEVER_MATCHED_STRING = "i_m_p_o_s_i_b_l_e_matcher";

    private TestWizard() {
    }

    public static String getExcelBDDStartPath(String childPath) {
        String absolutePath = new File(".").getAbsolutePath();
        return absolutePath.substring(0, absolutePath.lastIndexOf(childPath));
    }

    public static void showMap(Map<String, String> mapParams) {
        System.out.println(String.format("=======Header: %s=====", mapParams.get("Header")));
        for (Map.Entry<String, String> param : mapParams.entrySet()) {
            System.out.println(String.format("%s --- %s", param.getKey(), param.getValue()));
        }
    }

    public static int getInt(String string) {
        return Double.valueOf(string).intValue();
    }

    public static String makeMatcherString(String headerMatcher) {
        return ANY_MATCHER + headerMatcher + ANY_MATCHER;
    }
}
