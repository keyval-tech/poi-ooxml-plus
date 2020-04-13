package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteStringReplace;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class WriteStringReplaceProcessors implements WriteDataBodyProcessor {

    @Override
    public Object dataBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        WriteStringReplace writeStringReplace = (WriteStringReplace) annotation;

        String[] target = writeStringReplace.regex();
        String[] replacement = writeStringReplace.replacement();

        if (target.length != replacement.length || target.length == 0) {
            return columnValue;
        }

        String strValue = String.valueOf(columnValue);
        Map<String, Object> paramMap = new HashMap<>(2);
        paramMap.put("list", entityList);
        paramMap.put("i", entityListIndex);
        for (int i = 0; i < target.length; i++) {
            strValue = strValue
                    .replace(ElUtils.parseString(target[i], entityList, entityListIndex),
                            ElUtils.parseString(replacement[i], entityList, entityListIndex));
        }
        return strValue;
    }
}
