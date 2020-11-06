package com.kovizone.poi.ooxml.plus.command;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class WriteCellCommand extends WriteBaseCommand {
    /**
     * 列索引，当前行不存在列时为-1
     */
    private int cellIndex;
    /**
     * 默认最大列数
     */
    private int cellSize;

    /**
     * 默认列宽
     */
    private Integer defaultColumnWidth;

    /**
     * 优先输出空列的宽度集
     */
    private Integer[] blankCellWidths;

    /**
     * 创建单元格，
     * 注入样式
     *
     * @param styleName 样式名
     */
    public void createCell(String styleName) {
        if (blankCellWidths != null) {
            if (getCellIndex() < blankCellWidths.length) {
                for (int blankCellWidth : blankCellWidths) {
                    getRow().createCell(++cellIndex);
                    setCellWidth(blankCellWidth);
                }
            }
        }

        setCell(getRow().createCell(++cellIndex));
        if (defaultColumnWidth != null && defaultColumnWidth > 0) {
            setCellWidth(defaultColumnWidth);
        }
        CellStyle cellStyle = getStyleMap().get(styleName);
        if (cellStyle != null) {
            getCell().setCellStyle(cellStyle);
        }
    }

    /**
     * 设置当前列宽度，
     * 在新建Sheet或执行指令时渲染
     *
     * @param width 宽度
     */
    public void setCellWidth(int width) {
        int index = getCellIndex();
        if (width == -1 && defaultColumnWidth != null && defaultColumnWidth > 0) {
            width = defaultColumnWidth;
        }
        if (width > 0 && getLazyRenderCellWidth().get(index) == null) {
            getLazyRenderCellWidth().put(index, width);
        }
    }

    /**
     * 设置单元格样式
     *
     * @param styleName 样式名
     */
    public void setCellStyle(String styleName) {
        CellStyle cellStyle = getStyle(styleName);
        if (cellStyle != null) {
            getCell().setCellStyle(cellStyle);
        }
    }

    /**
     * 设置单元格值
     *
     * @param value 值
     */
    public void setCellValue(Object value) {
        if (value instanceof String) {
            getCell().setCellValue((String) value);
        }
        if (value instanceof Boolean) {
            getCell().setCellValue((Boolean) value);
        }
        if (value instanceof LocalDateTime) {
            getCell().setCellValue((LocalDateTime) value);
        }
        if (value instanceof LocalDate) {
            getCell().setCellValue((LocalDate) value);
        }
        if (value instanceof Date) {
            getCell().setCellValue((Date) value);
        }
        if (value instanceof Calendar) {
            getCell().setCellValue((Calendar) value);
        }
        if (value instanceof Byte) {
            getCell().setCellValue((Byte) value);
        }
        if (value instanceof Short) {
            getCell().setCellValue((Short) value);
        }
        if (value instanceof Integer) {
            getCell().setCellValue((Integer) value);
        }
        if (value instanceof Long) {
            getCell().setCellValue((Long) value);
        }
        if (value instanceof Float) {
            getCell().setCellValue((Float) value);
        }
        if (value instanceof Double) {
            getCell().setCellValue((Double) value);
        }
        if (value instanceof RichTextString) {
            getCell().setCellValue((RichTextString) value);
        }
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }

    public int getCellSize() {
        return cellSize;
    }

    public void setCellSize(int cellSize) {
        this.cellSize = cellSize;
    }

    public Integer getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    public void setDefaultColumnWidth(Integer defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
    }

    public Integer[] getBlankCellWidths() {
        return blankCellWidths;
    }

    public void setBlankCellWidths(Integer[] blankCellWidths) {
        this.blankCellWidths = blankCellWidths;
    }
}
