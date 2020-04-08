package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.DefaultValue;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * DefaultValue注解处理器
 *
 * @author KoviChen
 */
public class DefaultValueProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        if (targetFieldValue == null || "".equals(targetFieldValue.toString())) {
            DefaultValue defaultValue = (DefaultValue) annotation;
            return StringUtils.replaceFieldValue(defaultValue.value(), object);
        }
        return targetFieldValue;
    }
}
