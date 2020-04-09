package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.ColumnConfig;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteDateTitleProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * ColumnConfig注解处理器
 *
 * @author KoviChen
 */
public class ColumnConfigProcessors implements WriteDateTitleProcessor {

    @Override
    public void dateTitleProcess(Object annotation,
                                 ExcelCommand excelCommand,
                                 List<?> entityList,
                                 Field targetField) throws PoiOoxmlPlusException {
        ColumnConfig columnConfig = (ColumnConfig) annotation;
        String dateTitle = columnConfig.title();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = columnConfig.width();
        if (cellWidth >= 0) {
            excelCommand.setCellWidth(cellWidth);
        }
    }
}
