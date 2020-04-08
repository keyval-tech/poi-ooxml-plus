package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.WorkbookConstant;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

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

    /**
     * 去除多余空格
     *
     * @param arg 原字符串
     * @return 新字符串
     */
    public static String subExtraSpace(String arg) {
        return arg.replaceAll(" +", " ").trim();
    }

    public static String replaceFieldValue(String arg, Object object) throws PoiOoxmlPlusException {
        Map<String, String> paramMap = new HashMap<>();
        List<String> fieldNameList = FieldUtils.getFieldNameList(arg);
        for (String fieldName : fieldNameList) {
            try {
                Field field = object.getClass().getDeclaredField(fieldName);
                field.setAccessible(true);
                String node = String.valueOf(field.get(object));
                paramMap.put("#[" + fieldName + "]", node);
            } catch (NoSuchFieldException | IllegalAccessException e) {
                throw new PoiOoxmlPlusException("没有找到属性：" + fieldName);
            }
        }
        Set<Map.Entry<String, String>> entrySet = paramMap.entrySet();
        for (Map.Entry<String, String> entry : entrySet) {
            arg = arg.replace(entry.getKey(), entry.getValue());
        }
        return arg;
    }
}
