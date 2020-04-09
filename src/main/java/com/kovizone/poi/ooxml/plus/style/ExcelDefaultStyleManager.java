package com.kovizone.poi.ooxml.plus.style;

import com.kovizone.poi.ooxml.plus.ExcelHelper;
import com.kovizone.poi.ooxml.plus.command.ExcelStyleCommand;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * 默认样式
 *
 * @author KoviChen
 */
public class ExcelDefaultStyleManager implements ExcelStyleManager {

    @Override
    public Map<String, CellStyle> styleMap(ExcelStyleCommand command) {
        Map<String, CellStyle> styleMap = new HashMap<>();
        styleMap.put(ExcelHelper.HEADER_TITLE_CELL_STYLE_NAME, headerTitleCellStyle(command));
        styleMap.put(ExcelHelper.HEADER_TITLE_ROW_STYLE_NAME, headerTitleRowStyle(command));

        styleMap.put(ExcelHelper.HEADER_SUBTITLE_CELL_STYLE_NAME, headerSubtitleCellStyle(command));
        styleMap.put(ExcelHelper.HEADER_SUBTITLE_ROW_STYLE_NAME, headerSubtitleRowStyle(command));

        styleMap.put(ExcelHelper.DATE_TITLE_CELL_STYLE_NAME, dateTitleCellStyle(command));
        styleMap.put(ExcelHelper.DATE_TITLE_ROW_STYLE_NAME, dateTitleRowStyle(command));

        styleMap.put(ExcelHelper.DATE_BODY_CELL_STYLE_NAME, dateBodyCellStyle(command));
        styleMap.put(ExcelHelper.DATE_BODY_ROW_STYLE_NAME, dateBodyRowStyle(command));
        return styleMap;
    }

    private CellStyle headerTitleRowStyle(ExcelStyleCommand command) {
        return null;
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
        font.setColor(Font.COLOR_RED);
        font.setFontHeight((short) 500);
        cellStyle.setFont(font);
        return cellStyle;
    }

    private CellStyle headerSubtitleRowStyle(ExcelStyleCommand command) {
        return null;
    }

    private CellStyle headerSubtitleCellStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);

        Font cellFont = command.createFont();
        cellFont.setFontName("宋体");
        cellFont.setBold(true);
        cellFont.setFontHeight((short) 200);
        cellStyle.setFont(cellFont);
        return cellStyle;
    }

    private CellStyle dateTitleRowStyle(ExcelStyleCommand command) {
        return null;
    }

    private CellStyle dateTitleCellStyle(ExcelStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    private CellStyle dateBodyRowStyle(ExcelStyleCommand command) {
        return null;
    }

    private CellStyle dateBodyCellStyle(ExcelStyleCommand command) {
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
