package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.util.FieldUtils;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 属性替补处理器
 *
 * @author KoviChen
 */
public class SubstituteProcessors {

    /**
     * 处理工厂
     *
     * @param entity     实体
     * @param value      属性值
     * @param condition  替补条件
     * @param substitute 替补属性名
     * @return 替补值
     * @throws PoiOoxmlPlusException PoiPlusException
     */
    public static Object processorFactory(Object entity, Object value, String condition, String substitute) throws PoiOoxmlPlusException {
        if (condition.contains("IS NULL") || condition.contains("is null")) {
            return isNullProcessor(entity, value, substitute);
        }
        if (condition.contains("IS EMPTY") || condition.contains("is empty")) {
            return isEmptyProcessor(entity, value, substitute);
        }
        return null;
    }

    private static Object isEmptyProcessor(Object entity, Object value, String substitute) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        if (value == null || "".equals(value)) {
            return getValue(entity, substitute);
        }
        return value;
    }

    private static Object isNullProcessor(Object entity, Object value, String substitute) throws PoiOoxmlPlusException {
        if (value == null) {
            return getValue(entity, substitute);
        }
        return value;
    }

    /**
     * 解析出替补值
     */
    private static Object getValue(Object entity, String substitute) throws PoiOoxmlPlusException {
        Class<?> clazz = entity.getClass();
        List<String> substituteFieldNameList = FieldUtils.getFieldNameList(substitute);
        Map<String, String> substituteFieldValueMap = new HashMap<String, String>(16);
        for (String substituteFieldName : substituteFieldNameList) {
            try {
                Field conditionField = FieldUtils.getField(clazz, substituteFieldName);
                substituteFieldValueMap.put(substituteFieldName, (String) conditionField.get(entity));
            } catch (IllegalAccessException e) {
                throw new PoiOoxmlPlusException(e);
            }
        }
        Set<Map.Entry<String, String>> entrySet = substituteFieldValueMap.entrySet();
        for (Map.Entry<String, String> entry : entrySet) {
            substitute = substitute.replace("#[" + entry.getKey() + "]", entry.getValue());
        }
        return substitute;
    }
}
