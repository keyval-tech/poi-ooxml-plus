package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.HeaderRender;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.RegionUtil;

import java.lang.reflect.Field;
import java.util.List;

public class HeaderRenderProcessors implements WorkbookProcessor {


    @Override
    public void render(Object annotation, WorkbookCommand workbookCommand, Class<?> clazz, List<?> entityList, Field targetField) throws PoiOoxmlPlusException {
        HeaderRender headerRender = (HeaderRender) annotation;
        String headerTitle = headerRender.headerTitle();
        if (!StringUtils.isEmpty(headerTitle)) {
            Row row = workbookCommand.createRow();

            Cell cell = workbookCommand.createCell();
            cell.setCellValue(workbookCommand.textReplace(headerTitle));

            CellStyle cellStyle = workbookCommand.createCellStyle();
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);

            Font font = workbookCommand.createFont();
            font.setFontName("黑体");
            font.setBold(true);
            font.setColor(Font.COLOR_RED);
            font.setFontHeight((short) 500);
            cellStyle.setFont(font);

            cell.setCellStyle(cellStyle);
            workbookCommand.range(BorderStyle.THIN);

            short headerTitleHeight = headerRender.headerTitleHeight();
            if (headerTitleHeight != 0) {
                row.setHeight(headerTitleHeight);
            }
        }
        String headerSubtitle = headerRender.headerSubtitle();
        if (!StringUtils.isEmpty(headerSubtitle)) {
            Row row = workbookCommand.createRow();

            Cell cell = workbookCommand.createCell();
            cell.setCellValue(workbookCommand.textReplace(headerSubtitle));

            CellStyle cellStyle = workbookCommand.createCellStyle();

            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);

            Font cellFont = workbookCommand.createFont();
            cellFont.setFontName("宋体");
            cellFont.setBold(true);
            cellFont.setFontHeight((short) 200);
            cellStyle.setFont(cellFont);

            cell.setCellStyle(cellStyle);
            workbookCommand.range(BorderStyle.THIN);

            short headerSubtitleHeight = headerRender.headerSubtitleHeight();
            if (headerSubtitleHeight != 0) {
                row.setHeight(headerSubtitleHeight);
            }
        }

    }
}
