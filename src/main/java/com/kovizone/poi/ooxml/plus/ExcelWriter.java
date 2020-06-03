package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.anno.WriteColumnConfig;
import com.kovizone.poi.ooxml.plus.anno.WriteHeader;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
import com.kovizone.poi.ooxml.plus.processor.ProcessorFactory;
import com.kovizone.poi.ooxml.plus.style.ExcelDefaultStyle;
import com.kovizone.poi.ooxml.plus.util.ReflexUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.annotation.Resource;
import javax.xml.ws.soap.Addressing;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.*;

/**
 * POI构造者
 *
 * @author KoviChen
 * @author KoviChen2
 */
@Addressing
@Resource
@WriteHeader
public class ExcelWriter {

    /**
     * xlsz最大行数，默认{@value}
     */
    public static final int XLS_MAX_ROW_SIZE = 65536;

    /**
     * xlsx最大行数，默认{@value}
     */
    public static final int XLSX_MAX_ROW_SIZE = 1048576;

    /**
     * sheet码占位符
     */
    public static final String SHEET_NUM = "[page]";

    /**
     * 样式管理气
     */
    private ExcelStyle excelStyle;

    /**
     * 实体类构造器，
     * 注入默认样式
     */
    public ExcelWriter() {
        super();
        this.excelStyle = new ExcelDefaultStyle();
    }

    /**
     * 实体类构造器，
     * 注入自定义样式，
     * 自定义样式实现{@link ExcelStyle}
     */
    public ExcelWriter(ExcelStyle excelStyle) {
        super();
        this.excelStyle = excelStyle;
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList) throws ExcelWriteException {
        return writeSXSSF(entityList, null);
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList) throws ExcelWriteException {
        return writeXSSF(entityList, null);
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList) throws ExcelWriteException {
        return writeHSSF(entityList, null);
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws ExcelWriteException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws ExcelWriteException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
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
        write(workbook, entityList, null);
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

        if (entityList == null || entityList.isEmpty()) {
            return;
        }

        int maxLength = (workbook instanceof HSSFWorkbook) ? XLS_MAX_ROW_SIZE : XLSX_MAX_ROW_SIZE;

        // 主要属性集
        Class<?> clazz = entityList.get(0).getClass();
        List<Field> columnFieldList = columnFieldList(clazz);
        int cellSize = columnFieldList.size();
        ExcelCommand excelCommand = new ExcelCommand(workbook, cellSize, vars, excelStyle, entityList);

        sheetCycle:
        while (true) {
            excelCommand.createSheet();

            Annotation[] clazzAnnotations = ReflexUtils.getDeclaredAnnotations(clazz);
            for (Annotation clazzAnnotation : clazzAnnotations) {
                ProcessorFactory.sheetInitProcessor(clazzAnnotation, excelCommand, clazz);
                // headerProcessor子方法里创建行
                ProcessorFactory.headerProcessor(clazzAnnotation, excelCommand, clazz);
            }
            // 数据标题
            excelCommand.createRow();
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                excelCommand.createCell();

                for (Annotation clazzAnnotation : clazzAnnotations) {
                    ProcessorFactory.dateTitleProcessor(clazzAnnotation, excelCommand, field);
                }
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    ProcessorFactory.dateTitleProcessor(fieldAnnotation, excelCommand, field);
                }
            }

            // 数据集遍历
            for (; excelCommand.currentEntityListIndex() < entityList.size(); excelCommand.nextEntityListIndex()) {
                Object entity = entityList.get(excelCommand.currentEntityListIndex());

                if (excelCommand.currentRowIndex() >= maxLength) {
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
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    excelCommand.createCell();
                    for (Annotation clazzAnnotation : clazzAnnotations) {
                        value = ProcessorFactory.dateBodyProcessor(clazzAnnotation, excelCommand, field, value);
                    }
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        value = ProcessorFactory.dateBodyProcessor(fieldAnnotation, excelCommand, field, value);
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
     * 主要属性缓存
     */
    private static final Map<Class<?>, List<Field>> COLUMN_FIELD_LIST_CACHE = new HashMap<>(16);

    /**
     * 获取有@PoiColumn注解的属性集合，
     * 解析PoiColumn的sort，进行排序
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
                if (field.isAnnotationPresent(WriteColumnConfig.class)) {
                    WriteColumnConfig writeColumnConfig = field.getDeclaredAnnotation(WriteColumnConfig.class);
                    int sort = writeColumnConfig.sort();
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