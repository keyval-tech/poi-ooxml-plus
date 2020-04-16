package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteSubstitute;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * Substitute注解处理器
 *
 * @author KoviChen
 */
public class WriteSubstituteProcessors implements WriteDataBodyProcessor<WriteSubstitute> {

    @Override
    public Object dataBodyProcess(WriteSubstitute writeSubstitute,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        String criteria = writeSubstitute.criteria();
        String value = writeSubstitute.value();
        if (!StringUtils.isEmpty(value)) {
            if (ElUtils.parseBoolean(criteria, entityList, entityListIndex)) {
                return ElUtils.parseString(value, entityList, entityListIndex);
            }
        } else {
            return ElUtils.parseString(criteria, entityList, entityListIndex);
        }
        return columnValue;
    }
}
