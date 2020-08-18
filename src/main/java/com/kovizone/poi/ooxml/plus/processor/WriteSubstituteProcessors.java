package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteSubstitute;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * Substitute注解处理器
 *
 * @author KoviChen
 */
public class WriteSubstituteProcessors implements WriteRenderProcessor<WriteSubstitute> {

    @Override
    public Object dataBodyRender(WriteSubstitute writeSubstitute,
                                 ExcelCommand excelCommand,
                                 Field targetField,
                                 Object columnValue) {
        String criteria = writeSubstitute.criteria();
        String value = writeSubstitute.value();
        if (!StringUtils.isEmpty(value)) {
            if (excelCommand.parseBoolean(criteria)) {
                return excelCommand.parseString(value);
            }
        } else {
            return excelCommand.parseString(criteria);
        }
        return columnValue;
    }
}
