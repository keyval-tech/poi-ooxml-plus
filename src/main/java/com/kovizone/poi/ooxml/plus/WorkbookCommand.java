package com.kovizone.poi.ooxml.plus;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.util.*;

/**
 * 工作表基础命令
 *
 * @author KoviChen
 */
public class WorkbookCommand extends WorkbookStyleCommand {

    protected WorkbookCommand(Workbook workbook, int cellSize, Map<String, String> textReplaceMap) {
        super();
        this.workbook = workbook;
        this.cellSize = cellSize;
        this.textReplaceMap = textReplaceMap;
        this.sheetIndex = 0;
        this.nextCellIndex = 0;
        this.nextRowIndex = 0;
        this.lateRenderCellWidth = new HashMap<>();
    }

    /**
     * 工作簿
     */
    private Sheet sheet;

    /**
     * 行
     */
    private Row row;
    /**
     * 默认最大宽度
     */
    private int cellSize;
    /**
     * 下一工作簿索引
     */
    private int sheetIndex;
    /**
     * 下一列索引
     */
    private int nextCellIndex;
    /**
     * 下一行索引
     */
    private int nextRowIndex;
    /**
     * 表头文本替换
     */
    private Map<String, String> textReplaceMap;

    /**
     * 获取当前单元格索引
     *
     * @return 当前单元格索引
     */
    public int currentCowIndex() {
        return nextCellIndex - 1;
    }

    /**
     * 获取当前行索引
     *
     * @return 当前行索引
     */
    public int currentRowIndex() {
        return nextRowIndex - 1;
    }

    /**
     * 获取当前工作簿索引
     *
     * @return 当前工作簿索引
     */
    public int currentSheetIndex() {
        return sheetIndex - 1;
    }

    /**
     * 创建工作簿
     *
     * @param sheetName 工作簿名
     */
    public void createSheet(String sheetName) {
        if (!sheetName.contains(WorkbookConstant.SHEET_NUM)) {
            sheetName = sheetName + WorkbookConstant.SHEET_NUM;
        }
        Sheet sheet = workbook.createSheet(sheetName.replace(
                WorkbookConstant.SHEET_NUM,
                String.valueOf((sheetIndex) + 1)));
        this.sheet = sheet;

        // 重置列索引
        this.nextRowIndex = 0;
        // 重置行索引
        this.nextCellIndex = 0;
        if (sheetIndex > 1) {
            // 渲染之前的工作簿
            lateRender();
        }
        lateRenderFlag = false;
        sheetIndex++;
    }

    protected Sheet getSheet() {
        return sheet;
    }

    /**
     * 创建行
     *
     * @return 行
     */
    public Row createRow() {
        Row row = sheet.createRow(nextRowIndex++);
        this.row = row;
        // 重置行索引
        this.nextCellIndex = 0;
        return row;
    }

    /**
     * 创建行<BR/>
     * 注入样式
     *
     * @return 行
     */
    public Row createRow(CellStyle cellStyle) {
        Row row = createRow();
        if (cellStyle != null) {
            row.setRowStyle(cellStyle);
        }
        return row;
    }

    /**
     * 创建单元格
     *
     * @return 单元格
     */
    public Cell createCell() {
        Cell cell = row.createCell(nextCellIndex++);
        return cell;
    }

    /**
     * 创建单元格<BR/>
     * 注入样式
     *
     * @return 单元格
     */
    public Cell createCell(CellStyle cellStyle) {
        Cell cell = createCell();
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
        return cell;
    }

    /**
     * 设置列默认列宽<BR/>
     *
     * @param width 宽度
     */
    public void setDefaultColumnWidth(int width) {
        sheet.setDefaultColumnWidth(width);
    }

    /**
     * 设置列默认行高<BR/>
     *
     * @param height 高度
     */
    public void setDefaultRowHeight(short height) {
        sheet.setDefaultRowHeight(height);
    }

    private boolean lateRenderFlag;
    private Map<Integer, Integer> lateRenderCellWidth;

    protected void lateRender() {
        if (!lateRenderFlag) {
            lateRenderCellWidth.forEach(sheet::setColumnWidth);
            lateRenderFlag = true;
        }
    }

    /**
     * 设置当前列宽度<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     *
     * @param width 宽度
     */
    public void setCurrentCellWidth(int width) {
        lateRenderCellWidth.put(currentCowIndex(), width);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range() {
        range(null,
                null,
                null,
                null,
                null);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     *
     * @param cellRangeAddress 合并对象
     */
    public void range(CellRangeAddress cellRangeAddress) {
        range(cellRangeAddress,
                null,
                null,
                null,
                null);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(BorderStyle border) {
        range(null,
                border,
                border,
                border,
                border);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(BorderStyle borderTopAndBottom, BorderStyle borderLeftAndRight) {
        range(null,
                borderTopAndBottom,
                borderLeftAndRight,
                borderTopAndBottom,
                borderLeftAndRight);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(BorderStyle borderTop, BorderStyle borderRight, BorderStyle borderBottom, BorderStyle borderLeft) {
        range(null,
                borderTop,
                borderRight,
                borderBottom,
                borderLeft);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(CellRangeAddress region, BorderStyle border) {
        range(region,
                border,
                border,
                border,
                border);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(CellRangeAddress region, BorderStyle borderTopAndBottom, BorderStyle borderLeftAndRight) {
        range(region,
                borderTopAndBottom,
                borderLeftAndRight,
                borderTopAndBottom,
                borderLeftAndRight);
    }

    /**
     * 合并若干行<BR/>
     * 在新建Sheet或执行指令时渲染<BR/>
     */
    public void range(CellRangeAddress region, BorderStyle borderTop, BorderStyle borderRight, BorderStyle borderBottom, BorderStyle borderLeft) {
        if (region == null) {
            region = new CellRangeAddress(currentRowIndex(), currentRowIndex(), 0, cellSize);
        }
        sheet.addMergedRegionUnsafe(region);
        if (borderTop != null) {
            RegionUtil.setBorderTop(borderTop, region, sheet);
        }
        if (borderRight != null) {
            RegionUtil.setBorderRight(borderRight, region, sheet);
        }
        if (borderBottom != null) {
            RegionUtil.setBorderBottom(borderBottom, region, sheet);
        }
        if (borderLeft != null) {
            RegionUtil.setBorderLeft(borderLeft, region, sheet);
        }
    }

    public String textReplace(String target) {
        if (textReplaceMap == null) {
            return target;
        }
        Set<Map.Entry<String, String>> entrySet = textReplaceMap.entrySet();
        for (Map.Entry<String, String> entry : entrySet) {
            target = target.replace(entry.getKey(), entry.getValue());
        }
        return target;
    }

    public int getCellSize() {
        return cellSize;
    }
}
