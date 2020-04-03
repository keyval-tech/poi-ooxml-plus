package com.kovizone.poi.ooxml.plus.util;

/**
 * 字符串通用工具类
 *
 * @author KoviCHen
 */
public class StringUtils {

    public static boolean isEmpty(String arg) {
        if (arg == null || "".equals(arg.trim())) {
            return true;
        }
        return false;
    }

}
