package com.kovizone.poi.ooxml.plus.command;

import org.apache.poi.ss.usermodel.*;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
public class WriteBaseCommand {
    /**
     * 工作表
     */
    private Workbook workbook;

    /**
     * 当前工作簿
     */
    private Sheet sheet;
    /**
     * 当前行
     */
    private Row row;

    /**
     * 当前单元格
     */
    private Cell cell;
    /**
     * 样式管理
     */
    private Map<String, CellStyle> styleMap;
    /**
     * 文本替换
     */
    private Map<String, Object> replaceMap;
    /**
     * 当前Sheet是否已进行懒渲染
     */
    private boolean sheetLazyRenderFlag;
    /**
     * 当前Sheet需要进行懒渲染的列宽
     */
    private Map<Integer, Integer> lazyRenderCellWidth;
    /**
     * 当前Sheet需要进行懒渲染的行高
     */
    private Map<Integer, Short> lazyRenderRowHeight;

    public WriteBaseCommand() {
        sheetLazyRenderFlag = false;
    }

    /**
     * 获取样式
     *
     * @param styleName 样式名
     * @return 样式
     */
    public CellStyle getStyle(String styleName) {
        return getStyleMap().get(styleName);
    }

    /**
     * 文本替换
     *
     * @param target 原文
     * @return 替换后的文本
     * @see WriteBaseCommand#replaceMap
     */
    public String replace(String target) {
        if (replaceMap == null) {
            return target;
        }
        Set<Map.Entry<String, Object>> entrySet = replaceMap.entrySet();
        for (Map.Entry<String, Object> entry : entrySet) {
            target = target.replace(entry.getKey(), String.valueOf(entry.getValue()));
        }
        return target;
    }

    /**
     * 懒渲染
     */
    public void lazyRender() {
        lazyRender(sheet);
    }

    /**
     * 懒渲染
     */
    public void lazyRender(Sheet sheet) {
        if (!sheetLazyRenderFlag) {
            lazyRenderCellWidth.forEach((key, value) -> {
                if (value > 0) {
                    sheet.setColumnWidth(key, value);
                }
            });
            lazyRenderCellWidth = new HashMap<>(16);

            lazyRenderRowHeight.forEach((key, value) -> {
                sheet.getRow(key).setHeight(value);
            });
            lazyRenderRowHeight = new HashMap<>(16);

            sheetLazyRenderFlag = true;
        }
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Map<String, CellStyle> getStyleMap() {
        return styleMap;
    }

    public void setStyleMap(Map<String, CellStyle> styleMap) {
        this.styleMap = styleMap;
    }

    public Map<String, Object> getReplaceMap() {
        return replaceMap;
    }

    public void setReplaceMap(Map<String, Object> replaceMap) {
        this.replaceMap = replaceMap;
    }

    public Map<Integer, Integer> getLazyRenderCellWidth() {
        return lazyRenderCellWidth;
    }

    public void setLazyRenderCellWidth(Map<Integer, Integer> lazyRenderCellWidth) {
        if (lazyRenderCellWidth == null) {
            lazyRenderCellWidth = new HashMap<>(16);
        }
        this.lazyRenderCellWidth = lazyRenderCellWidth;
    }

    public Map<Integer, Short> getLazyRenderRowHeight() {
        return lazyRenderRowHeight;
    }

    public void setLazyRenderRowHeight(Map<Integer, Short> lazyRenderRowHeight) {
        if (lazyRenderRowHeight == null) {
            lazyRenderRowHeight = new HashMap<>(16);
        }
        this.lazyRenderRowHeight = lazyRenderRowHeight;
    }

    public boolean isSheetLazyRenderFlag() {
        return sheetLazyRenderFlag;
    }

    public void setSheetLazyRenderFlag(boolean sheetLazyRenderFlag) {
        this.sheetLazyRenderFlag = sheetLazyRenderFlag;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }
}
