package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.Substitute;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Substitute注解处理器
 *
 * @author KoviChen
 */
public class SubstituteProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        Substitute substitute = (Substitute) annotation;
        if (MvelUtils.parseBoolean(substitute.condition(), object)) {
            return StringUtils.replaceFieldValue(substitute.value(), object);
        }
        return targetFieldValue;
    }
}
