package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.util.POPUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel输入器
 *
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class ExcelReader {

    /**
     * 最大行号
     */
    private Integer maxRowSize;

    public <T> List<T> reader(Workbook workbook, Class<T> entityClass, Integer startRowIndex, Integer startCellIndex) {
        List<T> list = new ArrayList<>();

        List<Field> columnFieldList = POPUtils.columnFieldList(entityClass);
        int maxSheetNum = workbook.getNumberOfSheets();
        for (int s = 0; s < maxSheetNum; s++) {
            Sheet sheet = workbook.getSheetAt(s);
            int maxRowNum = sheet.getPhysicalNumberOfRows();

            if (maxRowNum > startRowIndex) {
                int maxCellNum = -1;
                for (int r = startRowIndex; r < maxRowNum; r++) {
                    Row row = sheet.getRow(r);
                    if (maxCellNum == -1) {
                        maxCellNum = row.getPhysicalNumberOfCells();
                    }
                    for (int c = startCellIndex; c < maxCellNum; c++) {
                        Cell cell = row.getCell(c);
                        CellType cellType = cell.getCellType();
                        switch (cellType) {
                            case _NONE:
                            case BLANK:
                                System.out.print("BLANK");
                                break;
                            case ERROR:
                                System.out.print("ERROR: " + cell.getErrorCellValue());
                                break;
                            case FORMULA:
                                System.out.print("FORMULA: " + cell.getCellFormula());
                                break;
                            case BOOLEAN:
                                System.out.print("BOOLEAN: " + cell.getBooleanCellValue());
                                break;
                            case NUMERIC:
                                System.out.print("NUMERIC: " + cell.getNumericCellValue());
                                break;
                            case STRING:
                                System.out.print("STRING: " + cell.getRichStringCellValue());
                                break;
                            default:
                                System.out.print("DEFAULT: " + cell.getStringCellValue());
                                break;
                        }
                        System.out.print(" | ");
                    }
                    System.out.println();
                }
            }
        }
        return list;
    }
}
