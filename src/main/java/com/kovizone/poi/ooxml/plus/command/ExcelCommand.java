package com.kovizone.poi.ooxml.plus.command;

import com.kovizone.poi.ooxml.plus.ExcelWriter;
import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.util.ElParser;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

/**
 * 工作表基础命令
 *
 * @author KoviChen
 */
public class ExcelCommand {

    /**
     * EL解析器
     */
    private ElParser elParser;
    /**
     * 数据遍历索引
     */
    private Integer entityListIndex;
    /**
     * 表达式解析器
     */
    private List<?> entityList;
    /**
     * 工作表
     */
    protected Workbook workbook;
    /**
     * 样式管理
     */
    private Map<String, CellStyle> styleMap;
    /**
     * 默认最大列数
     */
    private int cellSize;
    /**
     * 表头文本替换
     */
    private Map<String, Object> headerTextReplaceMap;

    /**
     * 默认Sheet标签名
     */
    private String defaultSheetName;

    /**
     * 优先输出空行的高度集
     */
    private Short[] blankRowHeights;

    /**
     * 优先输出空列的宽度集
     */
    private Integer[] blankCellWidths;

    /**
     * 默认列宽
     */
    private Integer defaultColumnWidth;

    /**
     * 默认行高
     */
    private Short defaultRowHeight;
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
     * 工作簿索引，当工作表不存在工作簿时为-1
     */
    private int sheetIndex;
    /**
     * 行索引，当工作簿不存在行时为-1
     */
    private int rowIndex;
    /**
     * 列索引，当前行不存在列时为-1
     */
    private int cellIndex;
    /**
     * 当前Sheet是否已进行懒渲染
     */
    private boolean sheetLazyRenderFlag;
    /**
     * 当前Sheet需要进行懒渲染的列宽
     */
    private Map<Integer, Integer> lateRenderCellWidth;
    /**
     * 当前Sheet需要进行懒渲染的行高
     */
    private Map<Integer, Short> lateRenderRowHeight;

    public ExcelCommand(
            Workbook workbook,
            ExcelStyle excelStyle,
            int cellSize,
            Map<String, Object> headerTextReplaceMap,
            List<?> entityList,
            ExcelInitCommand excelInitCommand) {
        super();
        this.elParser = new ElParser(entityList);
        this.workbook = workbook;
        this.cellSize = cellSize;
        this.headerTextReplaceMap = headerTextReplaceMap;
        this.entityList = entityList;

        this.defaultColumnWidth = excelInitCommand.getDefaultColumnWidth();
        this.defaultRowHeight = excelInitCommand.getDefaultRowHeight();
        this.blankRowHeights = excelInitCommand.getBlankRowHeights();
        this.blankCellWidths = excelInitCommand.getBlankCellWidths();
        this.defaultSheetName = excelInitCommand.getDefaultSheetName();

        this.styleMap = excelStyle.styleMap(new ExcelStyleCommand(workbook));

        this.sheetLazyRenderFlag = false;
        this.sheetIndex = -1;
        this.cellIndex = -1;
        this.rowIndex = -1;
        this.lateRenderCellWidth = new HashMap<>(16);
        this.lateRenderRowHeight = new HashMap<>(16);

        this.entityListIndex = 0;
    }

    /**
     * 获取当前单元格索引
     *
     * @return 当前单元格索引
     */
    public int currentCowIndex() {
        return cellIndex;
    }

    /**
     * 获取当前行索引
     *
     * @return 当前行索引
     */
    public int currentRowIndex() {
        return rowIndex;
    }

    /**
     * 获取当前工作簿索引
     *
     * @return 当前工作簿索引
     */
    public int currentSheetIndex() {
        return sheetIndex;
    }

    /**
     * 创建工作簿
     */
    public void createSheet() {
        createSheet(this.defaultSheetName);
    }

    /**
     * 创建工作簿
     *
     * @param sheetName 工作簿名
     */
    public void createSheet(String sheetName) {
        Sheet sheet;
        if (StringUtils.isEmpty(sheetName)) {
            sheetName = this.defaultSheetName;
        }
        if (!StringUtils.isEmpty(sheetName)) {
            if (!sheetName.contains(ExcelWriter.SHEET_NUM)) {
                sheetName = sheetName + "_" + ExcelWriter.SHEET_NUM;
            }
            sheet = workbook.createSheet(
                    sheetName.replace(ExcelWriter.SHEET_NUM, String.valueOf((++sheetIndex) + 1)));
        } else {
            sheet = workbook.createSheet("Sheet" + ((++sheetIndex) + 1));
        }
        this.sheet = sheet;

        // 重置列索引
        this.rowIndex = -1;
        // 重置行索引
        this.cellIndex = -1;
        if (sheetIndex > 0) {
            // 渲染之前的工作簿
            lazyRender();
        }
        sheetLazyRenderFlag = false;

        if (blankRowHeights != null) {
            for (Short blankRowHeight : blankRowHeights) {
                createRow(null);
                if (blankRowHeight > 0) {
                    setRowHeight(blankRowHeight);
                }
            }
        }
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
        row = sheet.createRow(++rowIndex);
        // 重置行索引
        this.cellIndex = -1;
        if (defaultRowHeight != null && defaultRowHeight > 0) {
            setRowHeight(defaultRowHeight);
        }

        if (blankCellWidths != null) {
            for (Integer blankCellWidth : blankCellWidths) {
                createCell(null);
                if (blankCellWidth > 0) {
                    setCellWidth(blankCellWidth);
                }
            }
        }
        CellStyle cellStyle = styleMap.get(styleName);
        if (cellStyle != null) {
            row.setRowStyle(cellStyle);
        }
    }

    /**
     * 创建单元格，
     * 注入样式
     *
     * @param styleName 样式名
     */
    public void createCell(String styleName) {
        this.cell = row.createCell(++cellIndex);
        if (defaultColumnWidth != null && defaultColumnWidth > 0) {
            setCellWidth(defaultColumnWidth);
        }
        CellStyle cellStyle = styleMap.get(styleName);
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 懒渲染
     */
    public void lazyRender() {
        if (!sheetLazyRenderFlag) {
            lateRenderCellWidth.forEach((key, value) -> {
                if (value > 0) {
                    sheet.setColumnWidth(key, value);
                }
            });
            lateRenderCellWidth = new HashMap<>(16);

            lateRenderRowHeight.forEach((key, value) -> {
                sheet.getRow(key).setHeight(value);
            });
            lateRenderRowHeight = new HashMap<>(16);

            sheetLazyRenderFlag = true;
        }
    }

    /**
     * 设置当前行高度
     *
     * @param height 高度
     */
    public void setRowHeight(short height) {
        int index = currentRowIndex();
        if (height == -1 && defaultRowHeight != null && defaultRowHeight > 0) {
            height = defaultRowHeight;
        }
        if (height > 0 && lateRenderRowHeight.get(index) == null) {
            lateRenderRowHeight.put(index, height);
        }
    }

    /**
     * 设置当前列宽度，
     * 在新建Sheet或执行指令时渲染
     *
     * @param width 宽度
     */
    public void setCellWidth(int width) {
        int index = currentCowIndex();
        if (width == -1 && defaultColumnWidth != null && defaultColumnWidth > 0) {
            width = defaultColumnWidth;
        }
        if (width > 0 && lateRenderCellWidth.get(index) == null) {
            lateRenderCellWidth.put(index, width);
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
                    currentRowIndex(),
                    currentRowIndex(),
                    blankRowSize,
                    cellSize - 1 + blankRowSize);
        }
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

    public String replace(String target) {
        if (headerTextReplaceMap == null) {
            return target;
        }
        Set<Map.Entry<String, Object>> entrySet = headerTextReplaceMap.entrySet();
        for (Map.Entry<String, Object> entry : entrySet) {
            target = target.replace(entry.getKey(), String.valueOf(entry.getValue()));
        }
        return target;
    }

    public int getCellSize() {
        return cellSize;
    }

    public void setCellValue(Object value) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        }
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }
        if (value instanceof LocalDateTime) {
            cell.setCellValue((LocalDateTime) value);
        }
        if (value instanceof LocalDate) {
            cell.setCellValue((LocalDate) value);
        }
        if (value instanceof Date) {
            cell.setCellValue((Date) value);
        }
        if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        }
        if (value instanceof Byte) {
            cell.setCellValue((Byte) value);
        }
        if (value instanceof Short) {
            cell.setCellValue((Short) value);
        }
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        }
        if (value instanceof Long) {
            cell.setCellValue((Long) value);
        }
        if (value instanceof Float) {
            cell.setCellValue((Float) value);
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
        }
        if (value instanceof RichTextString) {
            cell.setCellValue((RichTextString) value);
        }

    }

    public CellStyle getStyle(String styleName) {
        return styleMap.get(styleName);
    }

    public void setCellStyle(String styleName) {
        CellStyle cellStyle = getStyle(styleName);
        if (cellStyle != null) {
            this.cell.setCellStyle(cellStyle);
        }
    }

    public void setRowStyle(String styleName) {
        CellStyle cellStyle = getStyle(styleName);
        if (cellStyle != null) {
            this.row.setRowStyle(cellStyle);
        }
    }

    public <T> T parse(String expression, Class<T> clazz) {
        return elParser.parse(expression, entityListIndex, clazz);
    }

    public String parseString(String expression) {
        return elParser.parseString(expression, entityListIndex);
    }

    public Boolean parseBoolean(String expression) {
        return elParser.parseBoolean(expression, entityListIndex);
    }

    public List<?> getEntityList() {
        return entityList;
    }

    public int currentEntityListIndex() {
        return entityListIndex;
    }

    public void nextEntityListIndex() {
        if (entityListIndex == null) {
            entityListIndex = 0;
        } else {
            entityListIndex++;
        }
    }
}
