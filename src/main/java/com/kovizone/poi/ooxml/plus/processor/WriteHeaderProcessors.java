package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteHeader;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * HeaderRender注解处理器
 *
 * @author KoviChen
 */
public class WriteHeaderProcessors implements WriteRenderProcessor<WriteHeader> {

    /**
     * 表头数据行
     */
    public static final String HEADER_CELL_STYLE_NAME = "HEADER_TITLE_CELL_STYLE_NAME";

    @Override
    public void headerRender(WriteHeader writeHeader,
                             ExcelCommand excelCommand,
                             Class<?> clazz) {
        String headerTitle = writeHeader.value();
        if (!StringUtils.isEmpty(headerTitle)) {
            excelCommand.createCell(HEADER_CELL_STYLE_NAME);
            excelCommand.range(BorderStyle.THIN);
            excelCommand.setCellValue(excelCommand.replace(headerTitle));

            short headerTitleHeight = writeHeader.height();
            if (headerTitleHeight != 0) {
                excelCommand.setRowHeight(headerTitleHeight);
            }
        }
    }
}
