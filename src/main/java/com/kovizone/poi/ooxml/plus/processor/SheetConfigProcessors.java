package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.anno.SheetConfig;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.util.StringUtils;

import java.util.List;

/**
 * @author KoviChen
 */
public class SheetConfigProcessors implements WorkbookProcessor {

    @Override
    public void headerRender(Object annotation,
                             WorkbookCommand workbookCommand,
                             List<?> entityList,
                             Class<?> clazz) throws PoiOoxmlPlusException {
        SheetConfig sheetConfig = (SheetConfig) annotation;

        String sheetName = sheetConfig.sheetName();
        if (!StringUtils.isEmpty(sheetName)) {
            workbookCommand.setSheetName(sheetName.replace("#{sheetNum}", String.valueOf(workbookCommand.currentSheetIndex() + 1)));
        }

        int defaultColumnWidth = sheetConfig.defaultColumnWidth();
        if (defaultColumnWidth != 0) {
            workbookCommand.setDefaultColumnWidth(defaultColumnWidth);
        }
    }
}
