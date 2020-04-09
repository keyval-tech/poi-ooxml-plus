package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteStringReplace;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * StringReplace注解处理器
 *
 * @author KoviChen
 */
public class WriteStringReplaceAllProcessors implements WriteDateBodyProcessor {

    @Override
    public Object dateBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) throws PoiOoxmlPlusException {
        WriteStringReplace writeStringReplace = (WriteStringReplace) annotation;
        String[] regex = writeStringReplace.regex();
        String[] replacement = writeStringReplace.replacement();

        if (regex.length != replacement.length || regex.length == 0) {
            throw new PoiOoxmlPlusException("@StringReplaceAll配置有误");
        }

        String strValue = String.valueOf(columnValue);
        for (int i = 0; i < regex.length; i++) {
            strValue = strValue
                    .replaceAll(regex[i],
                            MvelUtils.parseString(replacement[i], entityList, entityListIndex));
        }

        return strValue;
    }
}
