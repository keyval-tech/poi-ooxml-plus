package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WritePrefix;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Prefix注解处理器
 *
 * @author KoviChen
 */
public class WritePrefixProcessors implements WriteDataBodyProcessor<WritePrefix> {

    @Override
    public Object dataBodyProcess(WritePrefix writePrefix,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        return ElUtils.parseString(writePrefix.value(), entityList, entityListIndex)
                .concat(String.valueOf(columnValue));
    }
}
