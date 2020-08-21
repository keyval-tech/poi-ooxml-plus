package com.kovizone.poi.ooxml.plus.command;

/**
 * 工作表基础命令
 *
 * @author KoviChen
 */
public class ExcelInitCommand {

    /**
     * 最大行号
     */
    private Integer maxRowSize;

    /**
     * 默认列宽
     */
    private Integer defaultColumnWidth;

    /**
     * 默认行高
     */
    private Short defaultRowHeight;


    /**
     * 优先输出空行的高度集
     */
    private Short[] blankRowHeights;

    /**
     * 优先输出空列的宽度集
     */
    private Integer[] blankCellWidths;

    /**
     * 默认Sheet标签名
     */
    private String defaultSheetName;

    public ExcelInitCommand() {
        this.maxRowSize = null;
        this.defaultColumnWidth = null;
        this.defaultRowHeight = null;
        this.blankRowHeights = null;
        this.blankCellWidths = null;
        this.defaultSheetName = null;
    }

    public String getDefaultSheetName() {
        return defaultSheetName;
    }

    public void setDefaultSheetName(String defaultSheetName) {
        this.defaultSheetName = defaultSheetName;
    }

    public Integer getMaxRowSize() {
        return maxRowSize;
    }

    public void setMaxRowSize(Integer maxRowSize) {
        this.maxRowSize = maxRowSize;
    }

    public Integer getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    public void setDefaultColumnWidth(Integer defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
    }

    public Short getDefaultRowHeight() {
        return defaultRowHeight;
    }

    public void setDefaultRowHeight(Short defaultRowHeight) {
        this.defaultRowHeight = defaultRowHeight;
    }

    public Short[] getBlankRowHeights() {
        return blankRowHeights;
    }

    public void setBlankRowHeights(Short[] blankRowHeights) {
        this.blankRowHeights = blankRowHeights;
    }

    public Integer[] getBlankCellWidths() {
        return blankCellWidths;
    }

    public void setBlankCellWidths(Integer[] blankCellWidths) {
        this.blankCellWidths = blankCellWidths;
    }
}
