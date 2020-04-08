package com.kovizone.poi.ooxml.plus.util;


import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 属性工具类
 *
 * @author KoviChen
 */
public class FieldUtils {

    /**
     * 属性名表达式正则表达式
     */
    private static final String FIELD_NAME_PATTERN = "(?<=\\#\\[)[^\\#\\[\\]]+(?=\\])";

    public static String getFieldName(String arg) {
        Matcher matcher = Pattern.compile(FIELD_NAME_PATTERN).matcher(arg);
        if (matcher.find()) {
            return matcher.group();
        }
        return null;
    }

    public static List<String> getFieldNameList(String arg) {
        Matcher matcher = Pattern.compile(FIELD_NAME_PATTERN).matcher(arg);
        List<String> fieldNames = new ArrayList<>();
        while (matcher.find()) {
            fieldNames.add(matcher.group());
        }
        return fieldNames;
    }

    public static Field getField(Class<?> clazz, String fieldName) throws PoiOoxmlPlusException {
        Field field = null;
        while (!clazz.equals(Object.class)) {
            try {
                field = clazz.getDeclaredField(fieldName);
                field.setAccessible(true);
                return field;
            } catch (NoSuchFieldException e) {
                clazz = clazz.getSuperclass();
            }
        }
        throw new PoiOoxmlPlusException("找不到属性：" + fieldName);
    }
}
