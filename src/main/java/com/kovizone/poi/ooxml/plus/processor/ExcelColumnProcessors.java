package com.kovizone.poi.ooxml.plus.processor;


import com.kovizone.poi.ooxml.plus.api.anno.ExcelColumn;
import com.kovizone.poi.ooxml.plus.api.processor.WriteRenderProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.lang.reflect.Field;

/**
 * <p>{@link ExcelColumn}注解的处理器</p>
 *
 * @author KoviChen
 */
public class ExcelColumnProcessors implements WriteRenderProcessor<ExcelColumn> {

    /**
     * 数据标题行
     */
    public static final String DATA_TITLE_CELL_STYLE_NAME = "DATA_TITLE_CELL_STYLE_NAME";

    /**
     * 数据体行
     */
    public static final String DATA_BODY_CELL_STYLE_NAME = "DATA_BODY_CELL_STYLE_NAME";

    @Override
    public void dataTitleRender(ExcelColumn excelColumn,
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
