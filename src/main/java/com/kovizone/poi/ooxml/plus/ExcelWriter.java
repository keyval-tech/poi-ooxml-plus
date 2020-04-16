package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.anno.WriteColumnConfig;
import com.kovizone.poi.ooxml.plus.anno.WriteHeader;
import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import com.kovizone.poi.ooxml.plus.api.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.api.style.ExcelStyle;
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
 */
@Addressing
@Resource
@WriteHeader
public class ExcelWriter {

    /**
     * xlsz最大行数
     */
    public static final int XLS_MAX_ROW_SIZE = 65536;

    /**
     * xlsx最大行数
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
     * 实体类构造器<BR/>
     * 注入默认样式
     */
    public ExcelWriter() {
        super();
        this.excelStyle = new ExcelDefaultStyle();
    }

    /**
     * 实体类构造器<BR/>
     * 注入自定义样式
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
        ExcelCommand excelCommand = new ExcelCommand(workbook, columnFieldList.size() - 1, vars, excelStyle);

        // 数据遍历索引
        int dateIndex = 0;

        sheetCycle:
        while (true) {
            excelCommand.createSheet();

            Annotation[] clazzAnnotations = ReflexUtils.getDeclaredAnnotations(clazz);
            for (Annotation clazzAnnotation : clazzAnnotations) {
                sheetInitProcessor(clazzAnnotation, excelCommand, clazz, entityList);
                // headerProcessor子方法里创建行
                headerProcessor(clazzAnnotation, excelCommand, clazz, entityList);
            }
            // 数据标题
            excelCommand.createRow();
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                excelCommand.createCell();

                for (Annotation clazzAnnotation : clazzAnnotations) {
                    dateTitleProcessor(clazzAnnotation, excelCommand, entityList, field);
                }
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    dateTitleProcessor(fieldAnnotation, excelCommand, entityList, field);
                }
            }

            // 数据集遍历
            for (; dateIndex < entityList.size(); dateIndex++) {
                Object entity = entityList.get(dateIndex);

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
                        e.printStackTrace();
                        throw new ExcelWriteException("读取属性值失败;" + e.getMessage());
                    }
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    excelCommand.createCell();
                    for (Annotation clazzAnnotation : clazzAnnotations) {
                        value = dateBodyProcessor(clazzAnnotation, excelCommand, dateIndex, entityList, field, value);
                    }
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        value = dateBodyProcessor(fieldAnnotation, excelCommand, dateIndex, entityList, field, value);
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

    @SuppressWarnings("unchecked")
    private <P> P getProcessor(Class<? extends Annotation> annotationClass, Class<P> processorClass) throws ExcelWriteException {
        // 判断注解是否存在处理器
        if (annotationClass.isAnnotationPresent(Processor.class)) {
            Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
            Class<?> processor = processorAnnotation.value();
            if (processorClass.isAssignableFrom(processor)) {
                try {
                    return (P) ReflexUtils.newInstance(processor);
                } catch (ReflexException e) {
                    e.printStackTrace();
                    throw new ExcelWriteException("构造处理器失败;" + e.getMessage());
                }
            }
        }
        return null;
    }

    /**
     * 解析表头处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param entityList   实体集
     * @throws ExcelWriteException 异常
     */
    private void sheetInitProcessor(Annotation annotation,
                                    ExcelCommand excelCommand,
                                    Class<?> clazz,
                                    List<?> entityList) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteSheetInitProcessor writeSheetInitProcessor = getProcessor(annotationClass, WriteSheetInitProcessor.class);
        if (writeSheetInitProcessor != null) {
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeSheetInitProcessor.sheetInitProcess((Annotation) annotationEntity,
                    excelCommand,
                    entityList,
                    clazz);
        }
    }

    /**
     * 解析表头处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param entityList   实体集
     * @throws ExcelWriteException 异常
     */
    private void headerProcessor(Annotation annotation,
                                 ExcelCommand excelCommand,
                                 Class<?> clazz,
                                 List<?> entityList) throws ExcelWriteException {

        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteHeaderProcessor writeHeaderProcessor = getProcessor(annotationClass, WriteHeaderProcessor.class);
        if (writeHeaderProcessor != null) {
            excelCommand.createRow();
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeHeaderProcessor.headerProcess((Annotation) annotationEntity,
                    excelCommand,
                    entityList,
                    clazz);
        }
    }

    /**
     * 解析数据标题处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param entityList   实体集
     * @param targetField  注解目标属性
     * @throws ExcelWriteException 异常
     */
    private void dateTitleProcessor(Annotation annotation,
                                    ExcelCommand excelCommand,
                                    List<?> entityList,
                                    Field targetField) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDataTitleProcessor writeDataTitleProcessor = getProcessor(annotationClass, WriteDataTitleProcessor.class);
        if (writeDataTitleProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            writeDataTitleProcessor.dataTitleProcess((Annotation) annotationEntity,
                    excelCommand,
                    entityList,
                    targetField);
        }
    }

    /**
     * 解析数据处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param entityList   实体集
     * @param targetField  注解目标属性
     * @param columnValue  注解目标属性值
     * @return 注解处理器更新后的值
     * @throws ExcelWriteException 异常
     */
    private Object dateBodyProcessor(Annotation annotation,
                                     ExcelCommand excelCommand,
                                     int entityListIndex,
                                     List<?> entityList,
                                     Field targetField,
                                     Object columnValue) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDataBodyProcessor writeDataBodyProcessor = getProcessor(annotationClass, WriteDataBodyProcessor.class);
        if (writeDataBodyProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            if (annotationEntity == null) {
                annotationEntity = entityList.get(entityListIndex).getClass().getDeclaredAnnotation(annotationClass);
            }
            columnValue = writeDataBodyProcessor.dataBodyProcess((Annotation) annotationEntity,
                    excelCommand,
                    entityList,
                    entityListIndex,
                    targetField,
                    columnValue);
        }
        return columnValue;
    }

    /**
     * 主要属性缓存
     */
    private static final Map<Class<?>, List<Field>> COLUMN_FIELD_LIST_CACHE = new HashMap<>(16);

    /**
     * 获取有@PoiColumn注解的属性集合<BR/>
     * 解析PoiColumn的sort，进行排序
     *
     * @param clazz 类
     * @return 获取有@PoiColumn注解的属性集合
     */
    private List<Field> columnFieldList(Class<?> clazz) throws ExcelWriteException {
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