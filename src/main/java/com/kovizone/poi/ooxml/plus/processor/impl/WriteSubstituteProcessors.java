package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteSubstitute;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Substitute注解处理器
 *
 * @author KoviChen
 */
public class WriteSubstituteProcessors implements WriteDataBodyProcessor {

    @Override
    public Object dataBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        WriteSubstitute writeSubstitute = (WriteSubstitute) annotation;
        Object object = entityList.get(entityListIndex);
        if (ElUtils.parseBoolean(writeSubstitute.criteria(), entityList,entityListIndex)) {
            return ElUtils.parseString(writeSubstitute.value(), entityList,entityListIndex);
        }
        return columnValue;
    }
}
