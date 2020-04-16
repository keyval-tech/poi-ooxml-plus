package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.anno.WriteColumnConfig;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * ColumnConfig注解处理器
 *
 * @author KoviChen
 */
public class WriteColumnConfigProcessors implements WriteDataTitleProcessor<WriteColumnConfig> {

    @Override
    public void dataTitleProcess(WriteColumnConfig writeColumnConfig,
                                 ExcelCommand excelCommand,
                                 List<?> entityList,
                                 Field targetField) {
        String dateTitle = writeColumnConfig.title();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = writeColumnConfig.width();
        if (cellWidth >= 0) {
            excelCommand.setCellWidth(cellWidth);
        }
    }
}
