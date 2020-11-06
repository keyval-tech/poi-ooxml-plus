package com.kovizone.poi.ooxml.plus.command;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class WriteRowCommand extends WriteCellCommand {
    /**
     * 行索引，当工作簿不存在行时为-1
     */
    private int rowIndex;

    /**
     * 默认行高
     */
    private Short defaultRowHeight;

    /**
     * 优先输出空行的高度集
     */
    private Short[] blankRowHeights;

    /**
     * 获取当前行索引
     *
     * @return 当前行索引
     */
    public int getRowIndex() {
        return rowIndex;
    }

    /**
     * 创建行
     */
    public void createRow() {
        createRow(null);
    }

    /**
     * 创建行，注入样式
     *
     * @param styleName 样式名
     */
    public void createRow(String styleName) {
        if (blankRowHeights != null) {
            if (getRowIndex() < blankRowHeights.length) {
                for (short blankRowHeight : blankRowHeights) {
                    getSheet().createRow(++rowIndex);
                    setRowHeight(blankRowHeight);
                }
            }
        }

        setRow(getSheet().createRow(++rowIndex));
        // 重置行索引
        super.setCellIndex(-1);
        setRowHeight((short) -1);

        CellStyle cellStyle = getStyleMap().get(styleName);
        if (cellStyle != null) {
            getRow().setRowStyle(cellStyle);
        }
    }

    /**
     * 设置当前行高度
     *
     * @param height 高度
     */
    public void setRowHeight(short height) {
        int index = getRowIndex();
        if (height == -1 && defaultRowHeight != null && defaultRowHeight > 0) {
            height = defaultRowHeight;
        }
        if (height > 0 && getLazyRenderRowHeight().get(index) == null) {
            getLazyRenderRowHeight().put(index, height);
        }
    }

    /**
     * 设置行样式
     *
     * @param styleName 样式名
     */
    public void setRowStyle(String styleName) {
        CellStyle cellStyle = getStyle(styleName);
        if (cellStyle != null) {
            getRow().setRowStyle(cellStyle);
        }
    }


    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     */
    public void range() {
        range(null,
                null,
                null,
                null,
                null);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染，
     *
     * @param cellRangeAddress 单元格范围地址
     */
    public void range(CellRangeAddress cellRangeAddress) {
        range(cellRangeAddress,
                null,
                null,
                null,
                null);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染，
     *
     * @param borderStyle 边框样式
     */
    public void range(BorderStyle borderStyle) {
        range(null,
                borderStyle,
                borderStyle,
                borderStyle,
                borderStyle);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     *
     * @param topAndBottomBorderStyle 上下边框样式
     * @param leftAndRightBorderStyle 左右边框样式
     */
    public void range(BorderStyle topAndBottomBorderStyle, BorderStyle leftAndRightBorderStyle) {
        range(null,
                topAndBottomBorderStyle,
                leftAndRightBorderStyle,
                topAndBottomBorderStyle,
                leftAndRightBorderStyle);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     *
     * @param topBorderStyle    上边框样式
     * @param rightBorderStyle  右边框样式
     * @param bottomBorderStyle 下边框样式
     * @param leftBorderStyle   左边框样式
     */
    public void range(BorderStyle topBorderStyle, BorderStyle rightBorderStyle, BorderStyle bottomBorderStyle, BorderStyle leftBorderStyle) {
        range(null,
                topBorderStyle,
                rightBorderStyle,
                bottomBorderStyle,
                leftBorderStyle);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     *
     * @param cellRangeAddress 单元格范围地址
     * @param borderStyle      边框样式
     */
    public void range(CellRangeAddress cellRangeAddress, BorderStyle borderStyle) {
        range(cellRangeAddress,
                borderStyle,
                borderStyle,
                borderStyle,
                borderStyle);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     *
     * @param cellRangeAddress        单元格范围地址
     * @param topAndBottomBorderStyle 上下边框样式
     * @param leftAndRightBorderStyle 左右边框样式
     */
    public void range(CellRangeAddress cellRangeAddress, BorderStyle topAndBottomBorderStyle, BorderStyle leftAndRightBorderStyle) {
        range(cellRangeAddress,
                topAndBottomBorderStyle,
                leftAndRightBorderStyle,
                topAndBottomBorderStyle,
                leftAndRightBorderStyle);
    }

    /**
     * 合并若干行，
     * 在新建Sheet或执行指令时渲染
     *
     * @param cellRangeAddress  单元格范围地址
     * @param topBorderStyle    上边框样式
     * @param rightBorderStyle  右边框样式
     * @param bottomBorderStyle 下边框样式
     * @param leftBorderStyle   左边框样式
     */
    public void range(CellRangeAddress cellRangeAddress, BorderStyle topBorderStyle, BorderStyle rightBorderStyle, BorderStyle bottomBorderStyle, BorderStyle leftBorderStyle) {
        if (cellRangeAddress == null) {
            int blankRowSize = (blankRowHeights == null ? 0 : blankRowHeights.length);

            cellRangeAddress = new CellRangeAddress(
                    getRowIndex(),
                    getRowIndex(),
                    blankRowSize,
                    getCellSize() - 1 + blankRowSize);
        }
        Sheet sheet = getSheet();
        sheet.addMergedRegionUnsafe(cellRangeAddress);
        if (topBorderStyle != null) {
            RegionUtil.setBorderTop(topBorderStyle, cellRangeAddress, sheet);
        }
        if (rightBorderStyle != null) {
            RegionUtil.setBorderRight(rightBorderStyle, cellRangeAddress, sheet);
        }
        if (bottomBorderStyle != null) {
            RegionUtil.setBorderBottom(bottomBorderStyle, cellRangeAddress, sheet);
        }
        if (leftBorderStyle != null) {
            RegionUtil.setBorderLeft(leftBorderStyle, cellRangeAddress, sheet);
        }
    }

    public Short getDefaultRowHeight() {
        return defaultRowHeight;
    }

    public void setDefaultRowHeight(Short defaultRowHeight) {
        this.defaultRowHeight = defaultRowHeight;
    }

    public WriteRowCommand() {
        rowIndex = -1;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public Short[] getBlankRowHeights() {
        return blankRowHeights;
    }

    public void setBlankRowHeights(Short[] blankRowHeights) {
        this.blankRowHeights = blankRowHeights;
    }
}
