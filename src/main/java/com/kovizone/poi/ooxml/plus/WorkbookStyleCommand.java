package com.kovizone.poi.ooxml.plus;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class WorkbookStyleCommand {
    /**
     * 工作表
     */
    protected Workbook workbook;

    /**
     * 创建单元格样式
     *
     * @return 单元格样式
     */
    public CellStyle createCellStyle() {
        return workbook.createCellStyle();
    }

    public Font createFont() {
        return workbook.createFont();
    }
}
