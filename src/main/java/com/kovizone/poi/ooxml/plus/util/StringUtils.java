package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import org.mvel2.MVEL;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

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

    public static String upperFirstCase(String arg) {
        char[] cs = arg.toCharArray();
        cs[0] -= 32;
        return String.valueOf(cs);
    }

}
