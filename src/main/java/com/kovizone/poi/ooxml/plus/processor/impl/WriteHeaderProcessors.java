package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteHeader;
import com.kovizone.poi.ooxml.plus.api.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.List;

/**
 * HeaderRender注解处理器
 *
 * @author KoviChen
 */
public class WriteHeaderProcessors implements WriteHeaderProcessor {

    @Override
    public void headerProcess(Object annotation,
                              ExcelCommand excelCommand,
                              List<?> entityList,
                              Class<?> clazz) {
        WriteHeader writeHeader = (WriteHeader) annotation;
        String headerTitle = writeHeader.value();
        if (!StringUtils.isEmpty(headerTitle)) {
            excelCommand.createCell(ExcelStyleCommand.HEADER_TITLE_CELL_STYLE_NAME);
            excelCommand.range(BorderStyle.THIN);
            excelCommand.setCellValue(excelCommand.eval(headerTitle));

            short headerTitleHeight = writeHeader.height();
            if (headerTitleHeight != 0) {
                excelCommand.getRow().setHeight(headerTitleHeight);
            }
        }
    }
}
