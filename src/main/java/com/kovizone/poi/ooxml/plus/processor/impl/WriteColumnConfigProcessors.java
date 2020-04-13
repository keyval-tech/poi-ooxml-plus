package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteColumnConfig;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * ColumnConfig注解处理器
 *
 * @author KoviChen
 */
public class WriteColumnConfigProcessors implements WriteDataTitleProcessor {

    @Override
    public void dataTitleProcess(Object annotation,
                                 ExcelCommand excelCommand,
                                 List<?> entityList,
                                 Field targetField) {
        WriteColumnConfig writeColumnConfig = (WriteColumnConfig) annotation;
        String dateTitle = writeColumnConfig.title();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = writeColumnConfig.width();
        if (cellWidth >= 0) {
            excelCommand.setCellWidth(cellWidth);
        }
    }
}
