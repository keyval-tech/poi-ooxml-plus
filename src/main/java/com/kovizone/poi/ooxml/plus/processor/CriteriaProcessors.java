package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.Criteria;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.util.FieldUtils;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.mvel2.MVEL;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 条件注解处理器
 *
 * @author KoviChen
 */
public class CriteriaProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation,
                             WorkbookCommand workbookCommand,
                             Object object,
                             List<?> entityList,
                             Class<?> clazz,
                             Field targetField,
                             Object targetFieldValue) throws PoiOoxmlPlusException {

        if (targetFieldValue == null) {
            return null;
        }
        Criteria criteria = (Criteria) annotation;
        if (MvelUtils.parseBoolean(criteria.value(), object)) {
            return targetFieldValue;
        }
        return null;
    }
}
