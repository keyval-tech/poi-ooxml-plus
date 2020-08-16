package com.kovizone.poi.ooxml.plus.api.style;

import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import com.kovizone.poi.ooxml.plus.processor.WriteHeaderProcessors;
import com.kovizone.poi.ooxml.plus.processor.WriteInitProcessors;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * 样式管理器接口
 *
 * @author KoviChen
 */
public interface ExcelStyle {

    /**
     * 默认表头样式
     *
     * @param command 样式创建指令
     * @return 样式
     */
    default CellStyle defaultHeaderStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        Font font = command.createFont();
        font.setFontName("黑体");
        font.setBold(true);
        font.setColor(Font.COLOR_NORMAL);
        font.setFontHeight((short) 500);
        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * 默认数据标题样式
     *
     * @param command 样式创建指令
     * @return 样式
     */
    default CellStyle defaultDataTitleStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }


    /**
     * 默认数据体样式
     *
     * @param command 样式创建指令
     * @return 样式
     */
    default CellStyle defaultDataBodyStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    /**
     * 单元格样式管理器
     *
     * @param command 样式创建指令
     * @return 样式Map
     */
    default Map<String, CellStyle> styleMap(ExcelStyleCommand command) {
        Map<String, CellStyle> styleMap = new HashMap<>();
        styleMap.put(WriteHeaderProcessors.HEADER_CELL_STYLE_NAME, defaultHeaderStyle(command));
        styleMap.put(WriteInitProcessors.DATA_TITLE_CELL_STYLE_NAME, defaultDataTitleStyle(command));
        styleMap.put(WriteInitProcessors.DATA_BODY_CELL_STYLE_NAME, defaultDataBodyStyle(command));
        return styleMap;
    }
}