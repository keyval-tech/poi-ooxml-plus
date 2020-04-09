package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.anno.ColumnConfig;
import com.kovizone.poi.ooxml.plus.anno.WriteHeaderRender;
import com.kovizone.poi.ooxml.plus.anno.base.Processor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.processor.WriteDateTitleProcessor;
import com.kovizone.poi.ooxml.plus.processor.WriteDateBodyProcessor;
import com.kovizone.poi.ooxml.plus.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.style.ExcelStyleManager;
import com.kovizone.poi.ooxml.plus.style.ExcelDefaultStyleManager;
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
@WriteHeaderRender
public class ExcelHelper {

    public static final String HEADER_TITLE_CELL_STYLE_NAME = "HEADER_TITLE_CELL_STYLE_NAME";
    public static final String HEADER_TITLE_ROW_STYLE_NAME = "HEADER_TITLE_ROW_STYLE_NAME";

    public static final String HEADER_SUBTITLE_CELL_STYLE_NAME = "HEADER_SUBTITLE_CELL_STYLE_NAME";
    public static final String HEADER_SUBTITLE_ROW_STYLE_NAME = "HEADER_SUBTITLE_ROW_STYLE_NAME";

    public static final String DATE_TITLE_CELL_STYLE_NAME = "DATE_TITLE_CELL_STYLE_NAME";
    public static final String DATE_TITLE_ROW_STYLE_NAME = "DATE_TITLE_ROW_STYLE_NAME";

    public static final String DATE_BODY_CELL_STYLE_NAME = "DATE_BODY_CELL_STYLE_NAME";
    public static final String DATE_BODY_ROW_STYLE_NAME = "DATE_BODY_ROW_STYLE_NAME";

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
    public static final String SHEET_NUM = "#{sheetNum}";

    /**
     * 样式管理气
     */
    private ExcelStyleManager excelStyleManager;

    /**
     * 实体类构造器<BR/>
     * 注入默认样式
     */
    public ExcelHelper() {
        super();
        this.excelStyleManager = new ExcelDefaultStyleManager();
    }

    /**
     * 实体类构造器<BR/>
     * 注入自定义样式
     */
    public ExcelHelper(ExcelStyleManager excelStyleManager) {
        super();
        this.excelStyleManager = excelStyleManager;
    }

    public static ExcelHelper getInstance() {
        return new ExcelHelper();
    }

    public static ExcelHelper getInstance(ExcelStyleManager excelStyleManager) {
        return new ExcelHelper(excelStyleManager);
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return writeSXSSF(entityList, null);
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeSXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new SXSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return writeXSSF(entityList, null);
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeXSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new XSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return writeHSSF(entityList, null);
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook writeHSSF(List<?> entityList, Map<String, Object> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new HSSFWorkbook();
        write(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList) throws PoiOoxmlPlusException {
        write(workbook, entityList, null);
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @param vars       替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public void write(Workbook workbook, List<?> entityList, Map<String, Object> vars) throws PoiOoxmlPlusException {

        if (entityList == null || entityList.isEmpty()) {
            return;
        }

        int maxLength = (workbook instanceof HSSFWorkbook) ? XLS_MAX_ROW_SIZE : XLSX_MAX_ROW_SIZE;

        // 主要属性集
        Class<?> clazz = entityList.get(0).getClass();
        List<Field> columnFieldList = columnFieldList(clazz);
        ExcelCommand excelCommand = new ExcelCommand(workbook, columnFieldList.size() - 1, vars, excelStyleManager);

        // 数据遍历索引
        int dateIndex = 0;

        sheetCycle:
        while (true) {
            excelCommand.createSheet();

            Annotation[] clazzAnnotations = ReflexUtils.getDeclaredAnnotations(clazz);
            for (Annotation clazzAnnotation : clazzAnnotations) {
                sheetInitProcessor(clazzAnnotation, excelCommand, clazz, entityList);
                // 不新增Sheet
                headerProcessor(clazzAnnotation, excelCommand, clazz, entityList);
            }
            // 数据标题
            excelCommand.createRow(DATE_TITLE_ROW_STYLE_NAME);
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                excelCommand.createCell(DATE_TITLE_CELL_STYLE_NAME);

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

                excelCommand.createRow(DATE_BODY_ROW_STYLE_NAME);
                for (Field field : columnFieldList) {
                    // 读取默认值
                    Object value = ReflexUtils.getValue(entity, field);
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    excelCommand.createCell(DATE_BODY_CELL_STYLE_NAME);
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

    private <P> P getProcessor(Class<? extends Annotation> annotationClass, Class<P> processorClass) throws PoiOoxmlPlusException {
        // 判断注解是否存在处理器
        if (annotationClass.isAnnotationPresent(Processor.class)) {
            Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
            Class<?> processor = processorAnnotation.value();
            if (processorClass.isAssignableFrom(processor)) {
                return ReflexUtils.newInstance(processorClass);
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
     * @throws PoiOoxmlPlusException 异常
     */
    private void sheetInitProcessor(Annotation annotation,
                                    ExcelCommand excelCommand,
                                    Class<?> clazz,
                                    List<?> entityList) throws PoiOoxmlPlusException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteSheetInitProcessor writeSheetInitProcessor = getProcessor(annotationClass, WriteSheetInitProcessor.class);
        if (writeSheetInitProcessor != null) {
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeSheetInitProcessor.sheetInitProcess(annotationEntity,
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
     * @throws PoiOoxmlPlusException 异常
     */
    private void headerProcessor(Annotation annotation,
                                 ExcelCommand excelCommand,
                                 Class<?> clazz,
                                 List<?> entityList) throws PoiOoxmlPlusException {

        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteHeaderProcessor writeHeaderProcessor = getProcessor(annotationClass, WriteHeaderProcessor.class);
        if (writeHeaderProcessor != null) {
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeHeaderProcessor.headerProcess(annotationEntity,
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
     * @throws PoiOoxmlPlusException 异常
     */
    private void dateTitleProcessor(Annotation annotation,
                                    ExcelCommand excelCommand,
                                    List<?> entityList,
                                    Field targetField) throws PoiOoxmlPlusException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDateTitleProcessor writeDateTitleProcessor = getProcessor(annotationClass, WriteDateTitleProcessor.class);
        if (writeDateTitleProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            writeDateTitleProcessor.dateTitleProcess(annotationEntity,
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
     * @throws PoiOoxmlPlusException 异常
     */
    private Object dateBodyProcessor(Annotation annotation,
                                     ExcelCommand excelCommand,
                                     int entityListIndex,
                                     List<?> entityList,
                                     Field targetField,
                                     Object columnValue) throws PoiOoxmlPlusException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDateBodyProcessor writeDateBodyProcessor = getProcessor(annotationClass, WriteDateBodyProcessor.class);
        if (writeDateBodyProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            if (annotationEntity == null) {
                annotationEntity = entityList.get(entityListIndex).getClass().getDeclaredAnnotation(annotationClass);
            }
            columnValue = writeDateBodyProcessor.dateBodyProcess(annotationEntity,
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
    private List<Field> columnFieldList(Class<?> clazz) throws PoiOoxmlPlusException {
        // 静态缓存
        List<Field> poiColumnFieldList = COLUMN_FIELD_LIST_CACHE.get(clazz);
        if (poiColumnFieldList != null) {
            return poiColumnFieldList;
        }

        List<Integer> sortList = new ArrayList<>(16);
        Map<Integer, Field> sortMap = new HashMap<>(16);

        while (!clazz.equals(Object.class)) {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                if (field.isAnnotationPresent(ColumnConfig.class)) {
                    ColumnConfig columnConfig = field.getDeclaredAnnotation(ColumnConfig.class);
                    int sort = columnConfig.sort();
                    if (sortMap.get(sort) != null) {
                        throw new PoiOoxmlPlusException("属性 " + field.getName() + " 的排序号被属性 " + sortMap.get(sort).getName() + " 占用：" + sort);
                    }
                    sortList.add(sort);
                    sortMap.put(sort, field);
                }
            }
            clazz = clazz.getSuperclass();
        }

        Collections.sort(sortList);
        poiColumnFieldList = new ArrayList<>();
        for (Integer sortNum : sortList) {
            poiColumnFieldList.add(sortMap.get(sortNum));
        }
        COLUMN_FIELD_LIST_CACHE.put(clazz, poiColumnFieldList);
        return poiColumnFieldList;
    }
}