package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.util.POPUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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
                        System.out.println(cell.getStringCellValue() + " | ");
                    }
                    System.out.println();
                }
            }
        }
        return list;
    }
}
