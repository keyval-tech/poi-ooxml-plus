package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteValue;
import com.kovizone.poi.ooxml.plus.api.processor.WriteProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * WriteValue注解处理器
 *
 * @author KoviChen
 */
public class WriteValueProcessors implements WriteProcessor<WriteValue> {

    @Override
    public Object dataBodyProcess(WriteValue writeValue,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        String[] values = writeValue.value();
        for (String value : values) {
            try {
                columnValue = excelCommand.parseString(value);
            } catch (Exception e) {
                if (!excelCommand.parseBoolean(value)) {
                    columnValue = null;
                }
            }
        }
        return columnValue;
    }
}