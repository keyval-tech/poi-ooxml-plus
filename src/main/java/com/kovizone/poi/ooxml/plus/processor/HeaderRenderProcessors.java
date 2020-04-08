package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.WorkbookCreator;
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
    public void headerRender(Object annotation,
                             WorkbookCommand workbookCommand,
                             List<?> entityList,
                             Class<?> clazz) throws PoiOoxmlPlusException {
        HeaderRender headerRender = (HeaderRender) annotation;
        String headerTitle = headerRender.headerTitle();
        if (!StringUtils.isEmpty(headerTitle)) {
            workbookCommand.createRow(WorkbookCreator.HEADER_TITLE_ROW_STYLE_NAME);
            workbookCommand.createCell(WorkbookCreator.HEADER_TITLE_CELL_STYLE_NAME);
            workbookCommand.range(BorderStyle.THIN);
            workbookCommand.setCellValue(workbookCommand.textReplace(headerTitle));

            short headerTitleHeight = headerRender.headerTitleHeight();
            if (headerTitleHeight != 0) {
                workbookCommand.setDefaultRowHeight(headerTitleHeight);
            }
        }
        String headerSubtitle = headerRender.headerSubtitle();
        if (!StringUtils.isEmpty(headerSubtitle)) {
            workbookCommand.createRow(WorkbookCreator.HEADER_SUBTITLE_ROW_STYLE_NAME);
            workbookCommand.createCell(WorkbookCreator.HEADER_SUBTITLE_CELL_STYLE_NAME);
            workbookCommand.range(BorderStyle.THIN);
            workbookCommand.setCellValue(workbookCommand.textReplace(headerSubtitle));

            short headerSubtitleHeight = headerRender.headerSubtitleHeight();
            if (headerSubtitleHeight != 0) {
                workbookCommand.setDefaultRowHeight(headerSubtitleHeight);
            }
        }
    }
}
