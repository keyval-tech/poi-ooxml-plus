package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.api.anno.ExcelColumn;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.processor.ExcelColumnProcessors;
import com.kovizone.poi.ooxml.plus.processor.ProcessorFactory;
import com.kovizone.poi.ooxml.plus.util.POPUtils;
import com.kovizone.poi.ooxml.plus.util.ReflexUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Excel输出器
 *
 * @author KoviChen
 */
public class ExcelWriter {

    /**
     * xls最大行数，默认{@value}
     */
    public static final int DEFAULT_XLS_MAX_ROW_SIZE = 65536;

    /**
     * xlsx最大行数，默认{@value}
     */
    public static final int DEFAULT_XLSX_MAX_ROW_SIZE = 1048576;

    /**
     * sheet码占位符
     */
    public static final String SHEET_NUM = "[page]";

    /**
     * 样式管理
     */
    private ExcelStyle excelStyle;
    /**
     * 默认行高
     */
    private Short defaultRowHeight;

    /**
     * 默认列宽
     */
    private Integer defaultColumnWidth;

    /**
     * 最大行号
     */
    private Integer maxRowSize;

    /**
     * 修改默认行高，创建新实例
     *
     * @param defaultRowHeight 默认行高
     * @return Excel输出器
     */
    public ExcelWriter setDefaultRowHeight(Short defaultRowHeight) {
        return new ExcelWriter(excelStyle, defaultRowHeight, defaultColumnWidth, maxRowSize);
    }

    /**
     * 修改默认列宽，创建新实例
     *
     * @param defaultColumnWidth 默认列宽
     * @return Excel输出器
     */
    public ExcelWriter setDefaultColumnWidth(Integer defaultColumnWidth) {
        return new ExcelWriter(excelStyle, defaultRowHeight, defaultColumnWidth, maxRowSize);
    }

    /**
     * 修改最大行号，创建新实例
     *
     * @param maxRowSize 最大行号
     * @return Excel输出器
     */
    public ExcelWriter setMaxRowSize(Integer maxRowSize) {
        return new ExcelWriter(excelStyle, defaultRowHeight, defaultColumnWidth, maxRowSize);
    }

    /**
     * 实体类构造器，
     * 注入默认样式
     */
    public ExcelWriter() {
        super();
        this.excelStyle = new ExcelStyle() {
        };
        this.defaultRowHeight = null;
        this.defaultColumnWidth = null;
        this.maxRowSize = null;
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     *
     * @param defaultRowHeight   默认行高
     * @param defaultColumnWidth 默认列宽
     */
    public ExcelWriter(Short defaultRowHeight, Integer defaultColumnWidth) {
        super();
        this.excelStyle = new ExcelStyle() {
        };
        this.defaultRowHeight = defaultRowHeight;
        this.defaultColumnWidth = defaultColumnWidth;
        this.maxRowSize = null;
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     *
     * @param excelStyle 样式
     */
    public ExcelWriter(ExcelStyle excelStyle) {
        super();
        this.excelStyle = excelStyle;
        this.defaultRowHeight = null;
        this.defaultColumnWidth = null;
        this.maxRowSize = null;
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     *
     * @param excelStyle         样式
     * @param defaultRowHeight   默认行高
     * @param defaultColumnWidth 默认列宽
     */
    public ExcelWriter(ExcelStyle excelStyle, Short defaultRowHeight, Integer defaultColumnWidth) {
        super();
        this.excelStyle = excelStyle;
        this.defaultRowHeight = defaultRowHeight;
        this.defaultColumnWidth = defaultColumnWidth;
        this.maxRowSize = null;
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     *
     * @param excelStyle 样式
     * @param maxRowSize 最大行数
     */
    public ExcelWriter(ExcelStyle excelStyle, Integer maxRowSize) {
        super();
        this.excelStyle = excelStyle;
        this.defaultRowHeight = null;
        this.defaultColumnWidth = null;
        this.maxRowSize = maxRowSize;
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     *
     * @param excelStyle         样式
     * @param defaultRowHeight   默认行高
     * @param defaultColumnWidth 默认列宽
     * @param maxRowSize         最大行数
     */
    public ExcelWriter(ExcelStyle excelStyle, Short defaultRowHeight, Integer defaultColumnWidth, Integer maxRowSize) {
        super();
        this.excelStyle = excelStyle;
        this.defaultRowHeight = defaultRowHeight;
        this.defaultColumnWidth = defaultColumnWidth;
        this.maxRowSize = maxRowSize;
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList) throws ExcelWriteException {
        return writeSXSSF(entityList, null, null);
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> replaceMap) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, replaceMap, null);
        return workbook;
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, String sheetName) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, null, sheetName);
        return workbook;
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> replaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, replaceMap, sheetName);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList) throws ExcelWriteException {
        return writeXSSF(entityList, null, null);
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> replaceMap) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, replaceMap);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, String sheetName) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, null, sheetName);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> replaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, replaceMap, sheetName);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList) throws ExcelWriteException {
        return writeHSSF(entityList, null, null);
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> replaceMap) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, replaceMap);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, String sheetName) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, null, sheetName);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @param sheetName  Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> replaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, replaceMap, sheetName);
        return workbook;
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList) throws ExcelWriteException {
        write(workbook, entityList, null, null);
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @param sheetName  Sheet标签名
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, String sheetName) throws ExcelWriteException {
        write(workbook, entityList, null, sheetName);
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, Map<String, Object> replaceMap) throws ExcelWriteException {
        write(workbook, entityList, replaceMap, null);
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @param replaceMap 替换文本
     * @param sheetName  Sheet标签名
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, Map<String, Object> replaceMap, String sheetName) throws ExcelWriteException {
        if (entityList == null || entityList.isEmpty()) {
            return;
        }

        int rowMaxLength = maxRowSize != null ? maxRowSize : ((workbook instanceof HSSFWorkbook) ? DEFAULT_XLS_MAX_ROW_SIZE : DEFAULT_XLSX_MAX_ROW_SIZE);

        // 主要属性集
        Class<?> clazz = entityList.get(0).getClass();
        List<Field> columnFieldList = POPUtils.columnFieldList(clazz);
        int cellSize = columnFieldList.size();
        ExcelCommand excelCommand = new ExcelCommand(workbook, cellSize, replaceMap, excelStyle, entityList, defaultRowHeight, defaultColumnWidth, sheetName);

        sheetCycle:
        while (true) {
            excelCommand.createSheet();

            Annotation[] clazzAnnotations = ReflexUtils.getDeclaredAnnotations(clazz);
            for (Annotation clazzAnnotation : clazzAnnotations) {
                excelCommand.createRow();
                ProcessorFactory.headerRender(clazzAnnotation, excelCommand, clazz);
            }

            // 数据标题
            excelCommand.createRow();
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                excelCommand.createCell(ExcelColumnProcessors.DATA_TITLE_CELL_STYLE_NAME);

                for (Annotation clazzAnnotation : clazzAnnotations) {
                    ProcessorFactory.dataTitleRender(clazzAnnotation, excelCommand, field);
                }
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    ProcessorFactory.dataTitleRender(fieldAnnotation, excelCommand, field);
                }
            }

            // 数据集遍历
            for (; excelCommand.currentEntityListIndex() < entityList.size(); excelCommand.nextEntityListIndex()) {
                Object entity = entityList.get(excelCommand.currentEntityListIndex());

                if (excelCommand.currentRowIndex() >= rowMaxLength) {
                    // 达到最大行数，新增工作簿
                    continue sheetCycle;
                }

                excelCommand.createRow();
                for (Field field : columnFieldList) {
                    // 读取默认值
                    Object value;
                    try {
                        value = ReflexUtils.getValue(entity, field);
                    } catch (ReflexException e) {
                        throw new ExcelWriteException("读取属性值失败;" + e.getMessage());
                    }
                    excelCommand.createCell(ExcelColumnProcessors.DATA_BODY_CELL_STYLE_NAME);
                    for (Annotation clazzAnnotation : clazzAnnotations) {
                        value = ProcessorFactory.dataBodyRender(clazzAnnotation, excelCommand, field, value);
                    }
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        value = ProcessorFactory.dataBodyRender(fieldAnnotation, excelCommand, field, value);
                    }
                    if (value != null) {
                        excelCommand.setCellValue(value);
                    }
                }
            }
            excelCommand.lateRender();
            break;
        }
    }
}