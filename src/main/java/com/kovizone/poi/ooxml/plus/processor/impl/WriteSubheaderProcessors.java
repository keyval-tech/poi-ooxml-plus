package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.anno.WriteSubheader;
import com.kovizone.poi.ooxml.plus.api.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;

import java.util.List;

/**
 * WriteSubheader注解处理器
 *
 * @author KoviChen
 */
public class WriteSubheaderProcessors implements WriteHeaderProcessor {

    @Override
    public void headerProcess(Object annotation,
                              ExcelCommand excelCommand,
                              List<?> entityList,
                              Class<?> clazz) {
        WriteSubheader writeSubheader = (WriteSubheader) annotation;
        String headerSubtitle = writeSubheader.value();
        if (!StringUtils.isEmpty(headerSubtitle)) {
            excelCommand.createCell(ExcelStyleCommand.HEADER_SUBTITLE_CELL_STYLE_NAME);
            excelCommand.range(BorderStyle.THIN);
            excelCommand.setCellValue(excelCommand.eval(headerSubtitle));

            short headerSubtitleHeight = writeSubheader.height();
            if (headerSubtitleHeight != 0) {
                excelCommand.getRow().setHeight(headerSubtitleHeight);
            }
        }
    }
}
