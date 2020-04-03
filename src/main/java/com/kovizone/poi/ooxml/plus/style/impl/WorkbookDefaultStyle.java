package com.kovizone.poi.ooxml.plus.style.impl;

import com.kovizone.poi.ooxml.plus.WorkbookStyleCommand;
import com.kovizone.poi.ooxml.plus.style.WorkbookStyle;
import org.apache.poi.ss.usermodel.*;

/**
 * 默认样式
 *
 * @author KoviChen
 */
public class WorkbookDefaultStyle implements WorkbookStyle {

    @Override
    public CellStyle headerTitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    @Override
    public CellStyle headerTitleCellStyle(WorkbookStyleCommand command) {
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

    @Override
    public CellStyle headerSubtitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    @Override
    public CellStyle headerSubtitleCellStyle(WorkbookStyleCommand command) {
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

    @Override
    public CellStyle dateTitleRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    @Override
    public CellStyle dateTitleCellStyle(WorkbookStyleCommand command) {
        CellStyle cellStyle = command.createCellStyle();

        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        return cellStyle;
    }

    @Override
    public CellStyle dateBodyRowStyle(WorkbookStyleCommand command) {
        return null;
    }

    @Override
    public CellStyle dateBodyCellStyle(WorkbookStyleCommand command) {
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
