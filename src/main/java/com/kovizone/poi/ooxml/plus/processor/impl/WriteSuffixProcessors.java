package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteSuffix;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Suffix注解处理器
 *
 * @author KoviChen
 */
public class WriteSuffixProcessors implements WriteDateBodyProcessor {

    @Override
    public Object dateBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) throws PoiOoxmlPlusException {
        WriteSuffix writeSuffix = (WriteSuffix) annotation;

        return String.valueOf(columnValue)
                .concat(MvelUtils.parseString(writeSuffix.value(), entityList,entityListIndex));
    }
}
