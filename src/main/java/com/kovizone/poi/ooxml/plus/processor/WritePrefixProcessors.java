package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WritePrefix;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * Prefix注解处理器
 *
 * @author KoviChen
 */
public class WritePrefixProcessors implements WriteDataBodyProcessor<WritePrefix> {

    @Override
    public Object dataBodyProcess(WritePrefix writePrefix,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        return excelCommand.parseString(writePrefix.value())
                .concat(String.valueOf(columnValue));
    }
}
