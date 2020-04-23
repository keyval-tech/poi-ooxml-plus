package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteInit;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * @author KoviChen
 */
public class WriteInitProcessors implements WriteSheetInitProcessor<WriteInit>, WriteDataBodyProcessor<WriteInit>, WriteDataTitleProcessor<WriteInit> {

    @Override
    public void sheetInitProcess(WriteInit writeInit,
                                 ExcelCommand excelCommand,
                                 Class<?> clazz) {

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
    public Object dataBodyProcess(WriteInit writeInit,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        excelCommand.setCellStyle(ExcelStyleCommand.DATA_BODY_CELL_STYLE_NAME);
        return columnValue;
    }

    @Override
    public void dataTitleProcess(WriteInit writeInit,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
        excelCommand.setCellStyle(ExcelStyleCommand.DATA_TITLE_CELL_STYLE_NAME);
    }
}
