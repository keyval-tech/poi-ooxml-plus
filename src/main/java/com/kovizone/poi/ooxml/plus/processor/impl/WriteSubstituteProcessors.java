package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteSubstitute;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.MvelUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Substitute注解处理器
 *
 * @author KoviChen
 */
public class WriteSubstituteProcessors implements WriteDateBodyProcessor {

    @Override
    public Object dateBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) throws PoiOoxmlPlusException {
        WriteSubstitute writeSubstitute = (WriteSubstitute) annotation;
        Object object = entityList.get(entityListIndex);
        if (MvelUtils.parseBoolean(writeSubstitute.criteria(), entityList,entityListIndex)) {
            return MvelUtils.parseString(writeSubstitute.value(), entityList,entityListIndex);
        }
        return columnValue;
    }
}
