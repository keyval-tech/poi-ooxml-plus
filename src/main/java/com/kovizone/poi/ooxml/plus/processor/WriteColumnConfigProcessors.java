package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.anno.WriteColumnConfig;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * <p>{@link WriteColumnConfig}注解的处理器</p>
 *
 * @author KoviChen
 */
public class WriteColumnConfigProcessors implements WriteDataTitleProcessor<WriteColumnConfig> {

    @Override
    public void dataTitleProcess(WriteColumnConfig writeColumnConfig,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
        String dateTitle = writeColumnConfig.value();
        if (!StringUtils.isEmpty(dateTitle)) {
            excelCommand.setCellValue(dateTitle);
        }
        int cellWidth = writeColumnConfig.width();
        excelCommand.setCellWidth(cellWidth);
    }
}
