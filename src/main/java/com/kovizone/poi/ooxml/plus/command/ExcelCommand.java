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

    public ExcelCommand(Workbook workbook, int cellSize, Map<String, Object> vars, ExcelStyle excelStyle, List<?> entityList) {
        super();
        this.workbook = workbook;
        this.cellSize = cellSize;
        this.vars = vars;
        this.sheetIndex = 0;
        this.nextCellIndex = 0;
        this.nextRowIndex = 0;
        this.lateRenderCellWidth = new HashMap<>(16);
        this.lateRenderRowHeight = new HashMap<>(16);
        this.styleMap = excelStyle.styleMap(new ExcelStyleCommand(workbook));
        this.entityList = entityList;
        this.elParser = new ElParser(entityList);
        this.entityListIndex = 0;
    }

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
     * 工作簿
     */
    private Sheet sheet;
    /**
     * 行
     */
    private Row row;
    /**
     * 单元格
     */
    private Cell cell;
    /**
     * 默认最大l列数
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
    private Map<String, Object> vars;

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
     */
    public void createSheet() {
        createSheet("Sheet#{sheetNum}");
    }

    /**
     * 创建工作簿
     *
     * @param sheetName 工作簿名
     */
    public void createSheet(String sheetName) {
        Sheet sheet;
        if (StringUtils.isEmpty(sheetName)) {
            if (!sheetName.contains(ExcelWriter.SHEET_NUM)) {
                sheetName = sheetName + ExcelWriter.SHEET_NUM;
            }
            sheet = workbook.createSheet(sheetName.replace(
                    ExcelWriter.SHEET_NUM,
                    String.valueOf((sheetIndex) + 1)));
        } else {
            sheet = workbook.createSheet();
        }
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
        setDefaultRowHeight(null);
    }

    public void setSheetName(String sheetName) {
        if (!sheetName.contains(ExcelWriter.SHEET_NUM)) {
            sheetName = sheetName + ExcelWriter.SHEET_NUM;
        }
        workbook.setSheetName(currentSheetIndex(), sheetName.replace(
                ExcelWriter.SHEET_NUM,
                String.valueOf(currentSheetIndex() + 1)));
    }

    /**
     * 获取Sheet
     *
     * @return sheet
     */
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
        if (defaultRowHeight != null) {
            setRowHeight(defaultRowHeight);
        }
        setDefaultColumnWidth(null);
        return row;
    }

    /**
     * 创建行，
     * 注入样式
     *
     * @param styleName 样式名
     * @return 行
     */
    public Row createRow(String styleName) {
        Row row = createRow();
        CellStyle cellStyle = styleMap.get(styleName);
        if (cellStyle != null) {
            row.setRowStyle(cellStyle);
        }
        return row;
    }

    public Row getRow() {
        return this.row;
    }

    /**
     * 创建单元格
     *
     * @return 单元格
     */
    public Cell createCell() {
        Cell cell = row.createCell(nextCellIndex++);
        this.cell = cell;
        if (defaultColumnWidth != null) {
            setCellWidth(defaultColumnWidth);
        }
        return cell;
    }

    /**
     * 创建单元格，
     * 注入样式
     *
     * @param styleName 样式名
     * @return 单元格
     */
    public Cell createCell(String styleName) {
        Cell cell = createCell();
        CellStyle cellStyle = styleMap.get(styleName);
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
        return cell;
    }

    public Cell getCell() {
        return this.cell;
    }

    private Integer defaultColumnWidth = null;

    private Short defaultRowHeight = null;

    /**
     * 设置列默认列宽，
     * 创建行后默认宽度会清除，
     *
     * @param width 宽度
     */
    public void setDefaultColumnWidth(Integer width) {
        defaultColumnWidth = width;
    }

    /**
     * 设置列默认行高，
     * 创建工作簿后默认高度会清除，
     *
     * @param height 高度
     */
    public void setDefaultRowHeight(Short height) {
        defaultRowHeight = height;
    }

    private boolean lateRenderFlag;
    private Map<Integer, Integer> lateRenderCellWidth;
    private Map<Integer, Short> lateRenderRowHeight;

    /**
     * 即时渲染列宽行高
     */
    public void lateRender() {
        if (!lateRenderFlag) {
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

            lateRenderFlag = true;
        }
    }

    /**
     * 设置当前行高度
     *
     * @param height 高度
     */
    public void setRowHeight(short height) {
        lateRenderRowHeight.put(currentRowIndex(), height);
    }

    /**
     * 设置当前列宽度，
     * 在新建Sheet或执行指令时渲染
     *
     * @param width 宽度
     */
    public void setCellWidth(int width) {
        if (width > 0) {
            lateRenderCellWidth.put(currentCowIndex(), width);
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
            cellRangeAddress = new CellRangeAddress(currentRowIndex(), currentRowIndex(), 0, cellSize - 1);
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

    public String eval(String target) {
        if (vars == null) {
            return target;
        }
        Set<Map.Entry<String, Object>> entrySet = vars.entrySet();
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

    public int nextEntityListIndex() {
        if (entityListIndex == null) {
            entityListIndex = 0;
        } else {
            entityListIndex++;
        }
        return entityListIndex;
    }

}
