package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteSuffix;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * Suffix注解处理器
 *
 * @author KoviChen
 */
public class WriteSuffixProcessors implements WriteRenderProcessor<WriteSuffix> {

    @Override
    public Object dataBodyRender(WriteSuffix writeSuffix,
                                 ExcelCommand excelCommand,
                                 Field targetField,
                                 Object columnValue) {
        return String.valueOf(columnValue)
                .concat(excelCommand.parseString(writeSuffix.value()));
    }
}
