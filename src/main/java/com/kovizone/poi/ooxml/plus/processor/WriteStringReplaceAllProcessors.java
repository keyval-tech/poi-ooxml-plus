package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteStringReplace;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class WriteStringReplaceAllProcessors implements WriteDataBodyProcessor<WriteStringReplace> {

    @Override
    public Object dataBodyProcess(WriteStringReplace writeStringReplace,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
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
                            ElUtils.parseString(replacement[i], entityList, entityListIndex));
        }

        return strValue;
    }
}
