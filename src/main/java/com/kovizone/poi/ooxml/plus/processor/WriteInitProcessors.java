package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteInit;
import com.kovizone.poi.ooxml.plus.api.processor.WriteProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * @author KoviChen
 */
public class WriteInitProcessors implements WriteProcessor<WriteInit> {

    /**
     * 数据标题行
     */
    public static final String DATA_TITLE_CELL_STYLE_NAME = "DATA_TITLE_CELL_STYLE_NAME";

    /**
     * 数据体行
     */
    public static final String DATA_BODY_CELL_STYLE_NAME = "DATA_BODY_CELL_STYLE_NAME";

    @Override
    public void headerProcess(WriteInit writeInit,
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
    public void dataTitleProcess(WriteInit writeInit,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
        excelCommand.setCellStyle(DATA_TITLE_CELL_STYLE_NAME);
    }

    @Override
    public Object dataBodyProcess(WriteInit writeInit,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        excelCommand.setCellStyle(DATA_BODY_CELL_STYLE_NAME);
        return columnValue;
    }
}
