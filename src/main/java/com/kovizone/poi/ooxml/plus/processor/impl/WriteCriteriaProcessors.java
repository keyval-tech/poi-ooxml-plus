package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteCriteria;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 条件注解处理器
 *
 * @author KoviChen
 */
public class WriteCriteriaProcessors implements WriteDataBodyProcessor {

    @Override
    public Object dataBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {

        if (columnValue == null) {
            return null;
        }
        WriteCriteria writeCriteria = (WriteCriteria) annotation;
        if (ElUtils.parseBoolean(writeCriteria.value(), entityList, entityListIndex)) {
            return columnValue;
        }
        return null;
    }
}