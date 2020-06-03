package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteStringReplace;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.reflect.Field;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class WriteStringReplaceProcessors implements WriteDataBodyProcessor<WriteStringReplace> {

    @Override
    public Object dataBodyProcess(WriteStringReplace writeStringReplace,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {

        String[] target = writeStringReplace.target();
        String[] replacement = writeStringReplace.replacement();

        if (target.length != replacement.length || target.length == 0) {
            return columnValue;
        }

        String strValue = String.valueOf(columnValue);
        for (int i = 0; i < target.length; i++) {
            String t = excelCommand.parseString(target[i]);
            String r = excelCommand.parseString(replacement[i]);
            strValue = strValue.replace(t, r);
        }
        return strValue;
    }
}
