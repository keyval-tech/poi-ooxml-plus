package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.FieldUtils;

import java.lang.reflect.Field;

/**
 * 单元格显示条件处理器
 *
 * @author KoviChen
 */
public class ConditionProcessors {

    public static Object processorFactory(Object entity, String condition) throws PoiOoxmlPlusException {
        if (condition.toUpperCase().contains("IS NOT NULL")) {
            return isNotNullProcessor(entity, condition);
        }
        if (condition.toUpperCase().contains("IS NULL")) {
            return isNullProcessor(entity, condition);
        }
        if (condition.toUpperCase().contains("IS EMPTY")) {
            return isEmptyProcessor(entity, condition);
        }
        if (condition.toUpperCase().contains("IS NOT EMPTY")) {
            return isNotEmptyProcessor(entity, condition);
        }
        return null;
    }

    private static Object isNotEmptyProcessor(Object entity, String condition) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        Object value = null;
        String conditionFieldName = FieldUtils.getFieldName(condition);
        try {
            Field conditionField = FieldUtils.getField(clazz, conditionFieldName);
            Object conditionFieldValue = conditionField.get(entity);
            if (conditionFieldValue == null || "".equals(conditionFieldValue)) {
                value = "";
            }
        } catch (IllegalAccessException e) {
            throw new PoiOoxmlPlusException(e);
        }
        return value;
    }

    private static Object isEmptyProcessor(Object entity, String condition) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        Object value = null;
        String conditionFieldName = FieldUtils.getFieldName(condition);
        try {
            Field conditionField = FieldUtils.getField(clazz, conditionFieldName);
            Object conditionFieldValue = conditionField.get(entity);
            if (conditionFieldValue != null && !"".equals(conditionFieldValue)) {
                value = "";
            }
        } catch (IllegalAccessException e) {
            throw new PoiOoxmlPlusException(e);
        }
        return value;
    }

    private static Object isNotNullProcessor(Object entity, String condition) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        Object value = null;
        String conditionFieldName = FieldUtils.getFieldName(condition);
        try {
            Field conditionField = FieldUtils.getField(clazz, conditionFieldName);
            Object conditionFieldValue = conditionField.get(entity);
            if (conditionFieldValue == null) {
                value = "";
            }
        } catch (IllegalAccessException e) {
            throw new PoiOoxmlPlusException(e);
        }
        return value;
    }

    private static Object isNullProcessor(Object entity, String condition) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        Object value = null;
        String conditionFieldName = FieldUtils.getFieldName(condition);
        try {
            Field conditionField = FieldUtils.getField(clazz, conditionFieldName);
            Object conditionFieldValue = conditionField.get(entity);
            if (conditionFieldValue != null) {
                value = "";
            }
        } catch (IllegalAccessException e) {
            throw new PoiOoxmlPlusException(e);
        }
        return value;


    }
}
