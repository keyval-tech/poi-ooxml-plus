package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteCriteria;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.util.ElUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 条件注解处理器
 *
 * @author KoviChen
 */
public class WriteCriteriaProcessors implements WriteDataBodyProcessor<WriteCriteria> {

    @Override
    public Object dataBodyProcess(WriteCriteria writeCriteria,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        try {
            if (ElUtils.parseBoolean(writeCriteria.value(), entityList, entityListIndex)) {
                return columnValue;
            }
            return null;
        } catch (Exception e) {
            String value = ElUtils.parseString(writeCriteria.value(), entityList, entityListIndex);
            return value;
        }
    }
}