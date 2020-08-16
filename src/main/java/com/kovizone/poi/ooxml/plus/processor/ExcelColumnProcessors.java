package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.api.anno.ExcelColumn;
import com.kovizone.poi.ooxml.plus.api.processor.WriteProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * <p>{@link ExcelColumn}注解的处理器</p>
 *
 * @author KoviChen
 */
public class ExcelColumnProcessors implements WriteProcessor<ExcelColumn> {

    @Override
    public void dataTitleProcess(ExcelColumn excelColumn,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
        String dateTitle = excelColumn.value();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = excelColumn.width();
        excelCommand.setCellWidth(cellWidth);
    }
}
