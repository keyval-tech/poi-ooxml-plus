package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.Prefix;
import com.kovizone.poi.ooxml.plus.anno.Suffix;
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
 * Suffix注解处理器
 *
 * @author KoviChen
 */
public class SuffixProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        Suffix suffix = (Suffix) annotation;
        return String.valueOf(targetFieldValue)
                .concat(StringUtils.replaceFieldValue(suffix.value(), clazz));
    }
}
