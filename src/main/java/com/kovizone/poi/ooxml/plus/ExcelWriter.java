package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.api.anno.ExcelColumn;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.processor.ExcelColumnProcessors;
import com.kovizone.poi.ooxml.plus.processor.ProcessorFactory;
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
     * 主要属性缓存
     */
    private static final Map<Class<?>, List<Field>> COLUMN_FIELD_LIST_CACHE = new HashMap<>(16);

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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap, null);
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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @param sheetName            Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap, sheetName);
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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @param sheetName            Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap, sheetName);
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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
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
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @param sheetName            Sheet标签名
     * @return 工作表
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap, String sheetName) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap, sheetName);
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
     * @param vars       替换文本
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, Map<String, Object> vars) throws ExcelWriteException {
        write(workbook, entityList, vars, null);
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @param vars       替换文本
     * @param sheetName  Sheet标签名
     * @throws ExcelWriteException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, Map<String, Object> vars, String sheetName) throws ExcelWriteException {
        if (entityList == null || entityList.isEmpty()) {
            return;
        }

        int rowMaxLength = maxRowSize != null ? maxRowSize : ((workbook instanceof HSSFWorkbook) ? DEFAULT_XLS_MAX_ROW_SIZE : DEFAULT_XLSX_MAX_ROW_SIZE);

        // 主要属性集
        Class<?> clazz = entityList.get(0).getClass();
        List<Field> columnFieldList = columnFieldList(clazz);
        int cellSize = columnFieldList.size();
        ExcelCommand excelCommand = new ExcelCommand(workbook, cellSize, vars, excelStyle, entityList, defaultRowHeight, defaultColumnWidth, sheetName);

        sheetCycle:
        while (true) {
            excelCommand.createSheet();

            Annotation[] clazzAnnotations = ReflexUtils.getDeclaredAnnotations(clazz);
            for (Annotation clazzAnnotation : clazzAnnotations) {
                // 方法内createRow
                ProcessorFactory.headerProcessor(clazzAnnotation, excelCommand, clazz);
            }

            // 数据标题
            excelCommand.createRow();
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                excelCommand.createCell(ExcelColumnProcessors.DATA_TITLE_CELL_STYLE_NAME);

                for (Annotation clazzAnnotation : clazzAnnotations) {
                    ProcessorFactory.dataTitleProcessor(clazzAnnotation, excelCommand, field);
                }
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    ProcessorFactory.dataTitleProcessor(fieldAnnotation, excelCommand, field);
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
                        value = ProcessorFactory.dataBodyProcessor(clazzAnnotation, excelCommand, field, value);
                    }
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        value = ProcessorFactory.dataBodyProcessor(fieldAnnotation, excelCommand, field, value);
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

    /**
     * 获取有{@link ExcelColumn}注解的属性集合，
     * 解析{@link ExcelColumn}的{@code sort}，进行排序
     *
     * @param clazz 类
     * @return 获取有@PoiColumn注解的属性集合
     */
    private List<Field> columnFieldList(Class<?> clazz) {
        // 静态缓存
        List<Field> poiColumnFieldList = COLUMN_FIELD_LIST_CACHE.get(clazz);
        if (poiColumnFieldList != null) {
            return poiColumnFieldList;
        }

        List<Integer> sortList = new ArrayList<>(16);
        Map<Field, Integer> sortMap = new HashMap<>(16);

        while (!clazz.equals(Object.class)) {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    ExcelColumn excelColumn = field.getDeclaredAnnotation(ExcelColumn.class);
                    int sort = excelColumn.sort();
                    if (!sortList.contains(sort)) {
                        sortList.add(sort);
                    }
                    sortMap.put(field, sort);
                }
            }
            clazz = clazz.getSuperclass();
        }

        Collections.sort(sortList);
        poiColumnFieldList = new ArrayList<>();
        for (Integer sortNum : sortList) {
            Set<Map.Entry<Field, Integer>> entrySet = sortMap.entrySet();
            for (Map.Entry<Field, Integer> entry : entrySet) {
                if (entry.getValue().equals(sortNum)) {
                    poiColumnFieldList.add(entry.getKey());
                }
            }
        }
        COLUMN_FIELD_LIST_CACHE.put(clazz, poiColumnFieldList);
        return poiColumnFieldList;
    }
}