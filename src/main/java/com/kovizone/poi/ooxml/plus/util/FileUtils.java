package com.kovizone.poi.ooxml.plus.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author KoviChen
 */
public class FileUtils {

    public static Workbook getWorkbook(String filePath) throws IOException {
        FileInputStream fileInputStream = null;
        try {
            Workbook workbook;
            try {
                fileInputStream = new FileInputStream(filePath);
                workbook = new HSSFWorkbook(fileInputStream);
            } catch (OfficeXmlFileException e) {
                if (fileInputStream != null) {
                    fileInputStream.close();
                }
                fileInputStream = new FileInputStream(filePath);
                workbook = new XSSFWorkbook(fileInputStream);
            }
            return workbook;
        } catch (Exception e) {
            if (fileInputStream != null) {
                fileInputStream.close();
            }
            throw e;
        }
    }

    public static Map<Integer, Integer> getColumnWidthList(Workbook workbook) {
        return getColumnWidthList(workbook, 0);
    }

    public static Map<Integer, Integer> getColumnWidthList(Workbook workbook, int sheetIndex) {
        return getColumnWidthList(workbook, sheetIndex, -1);
    }

    public static Map<Integer, Integer> getColumnWidthList(Workbook workbook, int sheetIndex, int maxCellSize) {
        if (sheetIndex < 0) {
            sheetIndex = 0;
        }

        Sheet sheet = workbook.getSheetAt(sheetIndex);
        Map<Integer, Integer> columnWidthMap = new HashMap<>();
        Row row = sheet.getRow(0);


        if (maxCellSize <= 0) {
            maxCellSize = row.getPhysicalNumberOfCells();
        }

        for (int i = 0; i < maxCellSize; i++) {
            columnWidthMap.put(i, sheet.getColumnWidth(i));
        }
        return columnWidthMap;
    }
}
