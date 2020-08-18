package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WritePrefix;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * Prefix注解处理器
 *
 * @author KoviChen
 */
public class WritePrefixProcessors implements WriteRenderProcessor<WritePrefix> {

    @Override
    public Object dataBodyRender(WritePrefix writePrefix,
                                 ExcelCommand excelCommand,
                                 Field targetField,
                                 Object columnValue) {
        return excelCommand.parseString(writePrefix.value())
                .concat(String.valueOf(columnValue));
    }
}
