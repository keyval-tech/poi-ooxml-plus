package com.kovizone.poi.ooxml.plus.processor.impl;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.SheetConfig;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;

import java.lang.reflect.Field;
import java.util.List;

/**
 * @author KoviChen
 */
public class SheetConfigProcessors implements WorkbookProcessor {

    @Override
    public Object render(Object annotation, WorkbookCommand workbookCommand, Class<?> clazz, List<?> entityList, Field targetField, Object value) throws PoiOoxmlPlusException {
        SheetConfig sheetConfig = (SheetConfig) annotation;
        String sheetName = sheetConfig.sheetName();
        workbookCommand.createSheet(sheetName);
        int defaultColumnWidth = sheetConfig.defaultColumnWidth();
        if (defaultColumnWidth != 0) {
            workbookCommand.setDefaultColumnWidth(defaultColumnWidth);
        }
        return value;
    }
}
