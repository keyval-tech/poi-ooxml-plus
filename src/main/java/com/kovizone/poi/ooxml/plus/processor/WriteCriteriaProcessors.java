package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteCriteria;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * 条件注解处理器
 *
 * @author KoviChen
 */
public class WriteCriteriaProcessors implements WriteDataBodyProcessor<WriteCriteria> {

    @Override
    public Object dataBodyProcess(WriteCriteria writeCriteria,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        try {
            if (excelCommand.parseBoolean(writeCriteria.value())) {
                return columnValue;
            }
            return null;
        } catch (Exception e) {
            String value = excelCommand.parseString(writeCriteria.value());
            return value;
        }
    }
}