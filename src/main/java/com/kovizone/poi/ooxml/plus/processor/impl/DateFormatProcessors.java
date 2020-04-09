package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.DateFormat;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * DateFormat注解处理器
 *
 * @author KoviChen
 */
public class DateFormatProcessors implements WriteDateBodyProcessor {

    @Override
    public Object dateBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) throws PoiOoxmlPlusException {
        if (!(columnValue instanceof Date)) {
            throw new PoiOoxmlPlusException("DateFormat注解的属性并非Date类型");
        }
        Date targetFieldDateValue = (Date) columnValue;
        DateFormat dateFormat = (DateFormat) annotation;
        if (StringUtils.isEmpty(dateFormat.value())) {
            return targetFieldDateValue.toString();
        }
        return new SimpleDateFormat(dateFormat.value()).format(targetFieldDateValue);
    }
}
