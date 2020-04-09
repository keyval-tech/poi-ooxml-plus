package com.kovizone.poi.ooxml.plus.command;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 工作表样式命令
 *
 * @author KoviChen
 */
public class ExcelStyleCommand {

    public static final String HEADER_TITLE_CELL_STYLE_NAME = "HEADER_TITLE_CELL_STYLE_NAME";
    public static final String HEADER_TITLE_ROW_STYLE_NAME = "HEADER_TITLE_ROW_STYLE_NAME";

    public static final String HEADER_SUBTITLE_CELL_STYLE_NAME = "HEADER_SUBTITLE_CELL_STYLE_NAME";
    public static final String HEADER_SUBTITLE_ROW_STYLE_NAME = "HEADER_SUBTITLE_ROW_STYLE_NAME";

    public static final String DATE_TITLE_CELL_STYLE_NAME = "DATE_TITLE_CELL_STYLE_NAME";
    public static final String DATE_TITLE_ROW_STYLE_NAME = "DATE_TITLE_ROW_STYLE_NAME";

    public static final String DATE_BODY_CELL_STYLE_NAME = "DATE_BODY_CELL_STYLE_NAME";
    public static final String DATE_BODY_ROW_STYLE_NAME = "DATE_BODY_ROW_STYLE_NAME";

    protected ExcelStyleCommand(Workbook workbook) {
        super();
        this.workbook = workbook;
    }

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

    /**
     * 创建字体样式
     *
     * @return 字体样式
     */
    public Font createFont() {
        return workbook.createFont();
    }
}
