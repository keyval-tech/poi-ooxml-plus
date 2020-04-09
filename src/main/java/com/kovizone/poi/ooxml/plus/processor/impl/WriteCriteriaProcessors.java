package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteCriteria;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 条件注解处理器
 *
 * @author KoviChen
 */
public class WriteCriteriaProcessors implements WriteDateBodyProcessor {

    @Override
    public Object dateBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) throws PoiOoxmlPlusException {

        if (columnValue == null) {
            return null;
        }
        WriteCriteria writeCriteria = (WriteCriteria) annotation;
        if (MvelUtils.parseBoolean(writeCriteria.value(), entityList, entityListIndex)) {
            return columnValue;
        }
        return null;
    }
}