package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.FieldUtils;
import com.kovizone.poi.ooxml.plus.anno.Prefix;
import com.kovizone.poi.ooxml.plus.anno.Suffix;

import java.lang.reflect.Field;

/**
 * 拼接处理器
 *
 * @author KoviChen
 */
public class ConcatProcessors {

    public static String concat(Object entity, Field field, String value) throws PoiOoxmlPlusException {
        if (field.isAnnotationPresent(Prefix.class)) {
            value = poiPrefixProcessor(entity, field) + value;
        }
        if (field.isAnnotationPresent(Suffix.class)) {
            value = value + poiSuffixProcessor(entity, field);
        }
        return value;
    }

    private static String poiPrefixProcessor(Object entity, Field field) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        Prefix prefix = field.getDeclaredAnnotation(Prefix.class);
        String[] prefixes = prefix.value();
        return getConcatString(entity, prefixes);
    }

    private static String poiSuffixProcessor(Object entity, Field field) throws PoiOoxmlPlusException {
        Suffix suffix = field.getDeclaredAnnotation(Suffix.class);
        String[] suffixes = suffix.value();
        return getConcatString(entity, suffixes);
    }

    private static String getConcatString(Object entity, String[] concats) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        StringBuilder suffixSb = new StringBuilder();
        for (String concat : concats) {
            String concatFieldString = "";
            String concatFieldName = FieldUtils.getFieldName(concat);
            if (concatFieldName != null) {
                try {
                    Field suffixField = FieldUtils.getField(clazz, concatFieldName);
                    Object suffixFieldValue = suffixField.get(entity);
                    concatFieldString = suffixFieldValue == null ? "" : String.valueOf(suffixFieldValue);
                } catch (IllegalAccessException e) {
                    throw new PoiOoxmlPlusException("读取属性失败：" + concatFieldName);
                }
            }
            suffixSb.append(concat.replace("#[" + concatFieldName + "]", concatFieldString));
        }
        return suffixSb.toString();
    }
}
