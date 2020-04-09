package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.ExcelHelper;
import com.kovizone.poi.ooxml.plus.anno.WriteHeaderRender;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * HeaderRender注解处理器
 *
 * @author KoviChen
 */
public class WriteHeaderProcessorProcessors implements WriteHeaderProcessor {

    @Override
    public void headerProcess(Object annotation,
                              ExcelCommand excelCommand,
                              List<?> entityList,
                              Class<?> clazz) throws PoiOoxmlPlusException {
        WriteHeaderRender writeHeaderRender = (WriteHeaderRender) annotation;
        String headerTitle = writeHeaderRender.headerTitle();
        if (!StringUtils.isEmpty(headerTitle)) {
            excelCommand.createRow(ExcelHelper.HEADER_TITLE_ROW_STYLE_NAME);
            excelCommand.createCell(ExcelHelper.HEADER_TITLE_CELL_STYLE_NAME);
            excelCommand.range(BorderStyle.THIN);
            excelCommand.setCellValue(excelCommand.eval(headerTitle));

            short headerTitleHeight = writeHeaderRender.headerTitleHeight();
            if (headerTitleHeight != 0) {
                excelCommand.getRow().setHeight(headerTitleHeight);
            }
        }
        String headerSubtitle = writeHeaderRender.headerSubtitle();
        if (!StringUtils.isEmpty(headerSubtitle)) {
            excelCommand.createRow(ExcelHelper.HEADER_SUBTITLE_ROW_STYLE_NAME);
            excelCommand.createCell(ExcelHelper.HEADER_SUBTITLE_CELL_STYLE_NAME);
            excelCommand.range(BorderStyle.THIN);
            excelCommand.setCellValue(excelCommand.eval(headerSubtitle));

            short headerSubtitleHeight = writeHeaderRender.headerSubtitleHeight();
            if (headerSubtitleHeight != 0) {
                excelCommand.getRow().setHeight(headerSubtitleHeight);
            }
        }
    }
}
