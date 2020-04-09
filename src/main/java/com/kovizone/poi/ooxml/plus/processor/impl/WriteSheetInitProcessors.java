package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteSheetInit;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.util.List;

/**
 * @author KoviChen
 */
public class WriteSheetInitProcessors implements WriteSheetInitProcessor {

    @Override
    public void sheetInitProcess(Object annotation,
                                 ExcelCommand excelCommand,
                                 List<?> entityList,
                                 Class<?> clazz) throws PoiOoxmlPlusException {
        WriteSheetInit writeSheetInit = (WriteSheetInit) annotation;

        String sheetName = writeSheetInit.sheetName();
        if (!StringUtils.isEmpty(sheetName)) {
            excelCommand.setSheetName(sheetName.replace("#{sheetNum}", String.valueOf(excelCommand.currentSheetIndex() + 1)));
        }

        int defaultColumnWidth = writeSheetInit.defaultColumnWidth();
        if (defaultColumnWidth >= 0) {
            excelCommand.setCellWidth(defaultColumnWidth);
        }

        short defaultRowHeight = writeSheetInit.defaultRowHeight();
        if (defaultRowHeight >= 0 && excelCommand.currentCowIndex() == 0) {
            excelCommand.setDefaultRowHeight(defaultRowHeight);
        }
    }
}
