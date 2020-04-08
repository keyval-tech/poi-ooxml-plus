package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.ColumnConfig;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * ColumnConfig注解处理器
 *
 * @author KoviChen
 */
public class ColumnConfigProcessors implements WorkbookProcessor {


    @Override
    public void titleRender(Object annotation, WorkbookCommand workbookCommand, List<?> entityList, Class<?> clazz, Field targetField) throws PoiOoxmlPlusException {
        ColumnConfig columnConfig = (ColumnConfig) annotation;
        String dateTitle = columnConfig.title();
        if (!StringUtils.isEmpty(dateTitle)) {
            workbookCommand.setCellValue(dateTitle);
        }
        int cellWidth = columnConfig.width();
        if (cellWidth >= 0) {
            workbookCommand.setCurrentCellWidth(cellWidth);
        }
    }
}
