package com.kovizone.poi.ooxml.plus.style;

import com.kovizone.poi.ooxml.plus.WorkbookCreator;
import com.kovizone.poi.ooxml.plus.WorkbookStyleCommand;
import com.kovizone.poi.ooxml.plus.style.WorkbookStyleManager;
import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;

/**
 * 默认样式
 *
 * @author KoviChen
 */
public class WorkbookDefaultStyleManager implements WorkbookStyleManager {

    @Override
    public Map<String, CellStyle> styleMap(WorkbookStyleCommand command) {
        Map<String, CellStyle> styleMap = new HashMap<>();
        styleMap.put(WorkbookCreator.HEADER_TITLE_CELL_STYLE_NAME, headerTitleCellStyle(command));
        styleMap.put(WorkbookCreator.HEADER_TITLE_ROW_STYLE_NAME, headerTitleRowStyle(command));

        styleMap.put(WorkbookCreator.HEADER_SUBTITLE_CELL_STYLE_NAME, headerSubtitleCellStyle(command));
        styleMap.put(WorkbookCreator.HEADER_SUBTITLE_ROW_STYLE_NAME, headerSubtitleRowStyle(command));

        styleMap.put(WorkbookCreator.DATE_TITLE_CELL_STYLE_NAME, dateTitleCellStyle(command));
        styleMap.put(WorkbookCreator.DATE_TITLE_ROW_STYLE_NAME, dateTitleRowStyle(command));

        styleMap.put(WorkbookCreator.DATE_BODY_CELL_STYLE_NAME, dateBodyCellStyle(command));
        styleMap.put(WorkbookCreator.DATE_BODY_ROW_STYLE_NAME, dateBodyRowStyle(command));
        return styleMap;
    }

    private CellStyle headerTitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    private CellStyle headerTitleCellStyle(WorkbookStyleCommand command) {
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

    private CellStyle headerSubtitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    private CellStyle headerSubtitleCellStyle(WorkbookStyleCommand command) {
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

    private CellStyle dateTitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    private CellStyle dateTitleCellStyle(WorkbookStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    private CellStyle dateBodyRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    private CellStyle dateBodyCellStyle(WorkbookStyleCommand command) {
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
