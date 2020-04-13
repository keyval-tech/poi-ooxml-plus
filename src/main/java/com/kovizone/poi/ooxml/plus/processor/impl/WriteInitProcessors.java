package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteInit;
import com.kovizone.poi.ooxml.plus.api.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author KoviChen
 */
public class WriteInitProcessors implements WriteSheetInitProcessor, WriteDataBodyProcessor, WriteDataTitleProcessor {

    @Override
    public void sheetInitProcess(Object annotation,
                                 ExcelCommand excelCommand,
                                 List<?> entityList,
                                 Class<?> clazz) {
        WriteInit writeInit = (WriteInit) annotation;

        String sheetName = writeInit.sheetName();
        if (!StringUtils.isEmpty(sheetName)) {
            excelCommand.setSheetName(sheetName.replace("#{sheetNum}", String.valueOf(excelCommand.currentSheetIndex() + 1)));
        }

        int defaultColumnWidth = writeInit.defaultColumnWidth();
        if (defaultColumnWidth >= 0) {
            excelCommand.setDefaultColumnWidth(defaultColumnWidth);
        }

        short defaultRowHeight = writeInit.defaultRowHeight();
        if (defaultRowHeight >= 0) {
            excelCommand.setDefaultRowHeight(defaultRowHeight);
        }
    }

    @Override
    public Object dataBodyProcess(Object annotation,
                                  ExcelCommand excelCommand,
                                  List<?> entityList,
                                  int entityListIndex,
                                  Field targetField,
                                  Object columnValue) {
        excelCommand.setCellStyle(ExcelStyleCommand.DATA_BODY_CELL_STYLE_NAME);
        return columnValue;
    }

    @Override
    public void dataTitleProcess(Object annotation, ExcelCommand excelCommand, List<?> entityList, Field targetField) {
        excelCommand.setCellStyle(ExcelStyleCommand.DATA_TITLE_CELL_STYLE_NAME);
    }
}
