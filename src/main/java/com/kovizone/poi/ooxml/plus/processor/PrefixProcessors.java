package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.Prefix;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.util.FieldUtils;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Prefix注解处理器
 *
 * @author KoviChen
 */
public class PrefixProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        Prefix prefix = (Prefix) annotation;
        return StringUtils.replaceFieldValue(prefix.value(), clazz)
                .concat(String.valueOf(targetFieldValue));
    }
}
