package com.kovizone.poi.ooxml.plus.util;

/**
 * 字符串通用工具类
 *
 * @author KoviCHen
 */
public class StringUtils {

    public static boolean isEmpty(String arg) {
        return arg == null || "".equals(arg.trim());
    }

    public static String upperFirstCase(String arg) {
        char[] cs = arg.toCharArray();
        cs[0] -= 32;
        return String.valueOf(cs);
    }

}
