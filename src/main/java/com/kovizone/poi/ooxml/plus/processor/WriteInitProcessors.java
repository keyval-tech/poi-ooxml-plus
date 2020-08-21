package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.WriteInit;
import com.kovizone.poi.ooxml.plus.api.processor.WriteInitProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelInitCommand;

/**
 * HeaderRender注解处理器
 *
 * @author KoviChen
 */
public class WriteInitProcessors implements WriteInitProcessor<WriteInit> {

    @Override
    public void init(WriteInit writeInit, ExcelInitCommand excelInitCommand, Class<?> clazz) {
        excelInitCommand.setDefaultSheetName(writeInit.sheetName());
        excelInitCommand.setDefaultRowHeight(writeInit.rowHeight());
        excelInitCommand.setDefaultColumnWidth(writeInit.columnWidth());
    }
}
