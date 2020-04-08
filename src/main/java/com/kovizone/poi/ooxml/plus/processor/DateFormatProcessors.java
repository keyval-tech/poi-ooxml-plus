package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.DateFormat;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import com.sun.corba.se.spi.orbutil.threadpool.Work;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * DateFormat注解处理器
 *
 * @author KoviChen
 */
public class DateFormatProcessors implements WorkbookProcessor {

    @Override
    public Object bodyRender(Object annotation, WorkbookCommand workbookCommand, Object object, List<?> entityList, Class<?> clazz, Field targetField, Object targetFieldValue) throws PoiOoxmlPlusException {
        if (!(targetFieldValue instanceof Date)) {
            throw new PoiOoxmlPlusException("DateFormat注解的属性并非Date类型");
        }
        Date targetFieldDateValue = (Date) targetFieldValue;
        DateFormat dateFormat = (DateFormat) annotation;
        if (StringUtils.isEmpty(dateFormat.value())) {
            return targetFieldDateValue.toString();
        }
        return new SimpleDateFormat(dateFormat.value()).format(targetFieldDateValue);
    }
}
