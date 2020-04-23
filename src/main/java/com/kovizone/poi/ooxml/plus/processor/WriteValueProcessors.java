package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteValue;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * WriteValue注解处理器
 *
 * @author KoviChen
 */
public class WriteValueProcessors implements WriteDataBodyProcessor<WriteValue> {

    @Override
    public Object dataBodyProcess(WriteValue writeCriteria,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        return excelCommand.parseString(writeCriteria.value());
    }
}