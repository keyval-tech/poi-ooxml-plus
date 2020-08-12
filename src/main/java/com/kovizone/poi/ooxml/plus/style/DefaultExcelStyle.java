package com.kovizone.poi.ooxml.plus.style;

import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * 默认样式
 *
 * @author KoviChen
 */
public class DefaultExcelStyle implements ExcelStyle {

    @Override
    public Map<String, CellStyle> styleMap(ExcelStyleCommand command) {
        Map<String, CellStyle> styleMap = new HashMap<>(4);
        styleMap.put(ExcelStyleCommand.HEADER_TITLE_CELL_STYLE_NAME, headerTitleCellStyle(command));
        styleMap.put(ExcelStyleCommand.DATA_TITLE_CELL_STYLE_NAME, dataTitleCellStyle(command));
        styleMap.put(ExcelStyleCommand.DATA_BODY_CELL_STYLE_NAME, dataBodyCellStyle(command));
        return styleMap;
    }

    private CellStyle headerTitleCellStyle(ExcelStyleCommand command) {
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

    private CellStyle dataTitleCellStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    private CellStyle dataBodyCellStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }
}
