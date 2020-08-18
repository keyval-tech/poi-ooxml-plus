package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteStringReplaceAll;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class WriteStringReplaceAllProcessors implements WriteRenderProcessor<WriteStringReplaceAll> {

    @Override
    public Object dataBodyRender(WriteStringReplaceAll writeStringReplace,
                                 ExcelCommand excelCommand,
                                 Field targetField,
                                 Object columnValue) {
        String[] regex = writeStringReplace.regex();
        String[] replacement = writeStringReplace.replacement();

        if (regex.length != replacement.length || regex.length == 0) {
            return columnValue;
        }

        String strValue = String.valueOf(columnValue);
        for (int i = 0; i < regex.length; i++) {
            strValue = strValue
                    .replaceAll(regex[i],
                            excelCommand.parseString(replacement[i]));
        }

        return strValue;
    }
}
