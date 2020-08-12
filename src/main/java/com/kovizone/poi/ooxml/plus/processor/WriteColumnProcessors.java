package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteColumn;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * <p>{@link WriteColumn}注解的处理器</p>
 *
 * @author KoviChen
 */
public class WriteColumnProcessors implements WriteDataTitleProcessor<WriteColumn> {

    @Override
    public void dataTitleProcess(WriteColumn writeColumn,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
        String dateTitle = writeColumn.value();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = writeColumn.width();
        excelCommand.setCellWidth(cellWidth);
    }
}
