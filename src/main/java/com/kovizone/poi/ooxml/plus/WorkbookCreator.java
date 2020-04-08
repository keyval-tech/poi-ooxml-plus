package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.anno.ColumnConfig;
import com.kovizone.poi.ooxml.plus.anno.HeaderRender;
import com.kovizone.poi.ooxml.plus.anno.Processor;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.style.WorkbookStyleManager;
import com.kovizone.poi.ooxml.plus.style.WorkbookDefaultStyleManager;
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
@HeaderRender
public class WorkbookCreator {

    public static final String HEADER_TITLE_CELL_STYLE_NAME = "HEADER_TITLE_CELL_STYLE_NAME";
    public static final String HEADER_TITLE_ROW_STYLE_NAME = "HEADER_TITLE_ROW_STYLE_NAME";

    public static final String HEADER_SUBTITLE_CELL_STYLE_NAME = "HEADER_SUBTITLE_CELL_STYLE_NAME";
    public static final String HEADER_SUBTITLE_ROW_STYLE_NAME = "HEADER_SUBTITLE_ROW_STYLE_NAME";

    public static final String DATE_TITLE_CELL_STYLE_NAME = "DATE_TITLE_CELL_STYLE_NAME";
    public static final String DATE_TITLE_ROW_STYLE_NAME = "DATE_TITLE_ROW_STYLE_NAME";

    public static final String DATE_BODY_CELL_STYLE_NAME = "DATE_BODY_CELL_STYLE_NAME";
    public static final String DATE_BODY_ROW_STYLE_NAME = "DATE_BODY_ROW_STYLE_NAME";

    /**
     * 工作表命令
     */
    private WorkbookCommand workbookCommand;

    /**
     * 样式管理气
     */
    WorkbookStyleManager workbookStyleManager;

    /**
     * 样式Map
     */
    Map<String, CellStyle> styleMap;

    /**
     * 实体类构造器<BR/>
     * 注入默认样式
     */
    public WorkbookCreator() {
        super();
        this.workbookStyleManager = new WorkbookDefaultStyleManager();
    }

    /**
     * 实体类构造器<BR/>
     * 注入自定义样式
     */
    public WorkbookCreator(WorkbookStyleManager workbookStyleManager) {
        super();
        this.workbookStyleManager = workbookStyleManager;
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorSXSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return constructorSXSSF(entityList, null);
    }

    /**
     * 构造SXSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorSXSSF(List<?> entityList, Map<String, String> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new SXSSFWorkbook();
        constructor(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorXSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return constructorXSSF(entityList, null);
    }

    /**
     * 构造XSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorXSSF(List<?> entityList, Map<String, String> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new XSSFWorkbook();
        constructor(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorHSSF(List<?> entityList) throws PoiOoxmlPlusException {
        return constructorHSSF(entityList, null);
    }

    /**
     * 构造HSSF工作表
     *
     * @param entityList           实体对象集
     * @param headerTextReplaceMap 表头替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public Workbook constructorHSSF(List<?> entityList, Map<String, String> headerTextReplaceMap) throws PoiOoxmlPlusException {
        Workbook workbook = new HSSFWorkbook();
        constructor(workbook, entityList, headerTextReplaceMap);
        return workbook;
    }

    /**
     * 构造工作表
     *
     * @param workbook   工作表
     * @param entityList 实体对象集
     * @throws PoiOoxmlPlusException 构造异常
     */
    public void constructor(Workbook workbook, List<?> entityList) throws PoiOoxmlPlusException {
        constructor(workbook, entityList, null);
    }


    /**
     * 构造工作表
     *
     * @param workbook       工作表
     * @param entityList     实体对象集
     * @param textReplaceMap 替换文本
     * @throws PoiOoxmlPlusException 构造异常
     */
    public void constructor(Workbook workbook, List<?> entityList, Map<String, String> textReplaceMap) throws PoiOoxmlPlusException {

        if (entityList == null || entityList.isEmpty()) {
            return;
        }

        Class<?> clazz = entityList.get(0).getClass();

        int maxLength;
        if (workbook instanceof HSSFWorkbook) {
            maxLength = WorkbookConstant.XLS_MAX_ROW_SIZE;
        } else {
            maxLength = WorkbookConstant.XLSX_MAX_ROW_SIZE;
        }

        // 生成内容
        List<Field> columnFieldList = poiColumnFieldList(clazz);
        workbookCommand = new WorkbookCommand(workbook, columnFieldList.size() - 1, textReplaceMap, workbookStyleManager);

        // 数据遍历索引
        int dateIndex = 0;

        sheetCycle:
        while (true) {
            workbookCommand.createSheet();

            Annotation[] clazzAnnotations = clazz.getDeclaredAnnotations();
            for (Annotation clazzAnnotation : clazzAnnotations) {
                // 不新增Sheet
                headerRender(clazzAnnotation, workbookCommand, clazz, entityList);
            }
            // 数据标题
            workbookCommand.createRow(DATE_TITLE_ROW_STYLE_NAME);
            for (Field field : columnFieldList) {
                Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                workbookCommand.createCell(DATE_TITLE_CELL_STYLE_NAME);
                for (Annotation fieldAnnotation : fieldAnnotations) {
                    titleRender(fieldAnnotation, workbookCommand, clazz, entityList, field);
                }
            }

            // 数据集遍历
            for (; dateIndex < entityList.size(); dateIndex++) {
                Object entity = entityList.get(dateIndex);

                if (workbookCommand.currentRowIndex() >= maxLength) {
                    // 达到最大行数，新增工作簿
                    continue sheetCycle;
                }

                workbookCommand.createRow(DATE_BODY_ROW_STYLE_NAME);
                for (Field field : columnFieldList) {
                    // 读取默认值
                    Object value = null;
                    try {
                        value = field.get(entity);
                    } catch (IllegalAccessException e) {
                        throw new PoiOoxmlPlusException("读取值失败");
                    }
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    workbookCommand.createCell(DATE_BODY_CELL_STYLE_NAME);
                    for (Annotation fieldAnnotation : fieldAnnotations) {
                        value = bodyRender(fieldAnnotation, workbookCommand, clazz, entity, entityList, field, value);
                        if (value != null) {
                            workbookCommand.setCellValue(value);
                        }
                    }
                }
            }
            workbookCommand.lateRender();
            break;
        }
    }

    /**
     * 解析表头处理器
     *
     * @param annotation      注解
     * @param workbookCommand 基础命令
     * @param entityList      实体集
     * @throws PoiOoxmlPlusException 异常
     */
    private void headerRender(Annotation annotation,
                              WorkbookCommand workbookCommand,
                              Class<?> clazz,
                              List<?> entityList) throws PoiOoxmlPlusException {

        Class<? extends Annotation> annotationClass = annotation.annotationType();
        // 判断注解是否存在处理器
        if (annotationClass.isAnnotationPresent(Processor.class)) {
            Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);

            // 处理器必须实现WorkbookProcessor
            Class<?> processor = processorAnnotation.value();
            if (WorkbookProcessor.class.isAssignableFrom(processor)) {
                try {
                    WorkbookProcessor workbookProcessor = (WorkbookProcessor) processor.newInstance();
                    workbookProcessor.headerRender(annotationEntity, workbookCommand, entityList, clazz);
                } catch (IllegalAccessException | InstantiationException e) {
                    throw new PoiOoxmlPlusException("注解处理器实例化失败;".concat(e.getMessage()));
                }
            }
        }
    }

    /**
     * 解析数据标题处理器
     *
     * @param annotation      注解
     * @param workbookCommand 基础命令
     * @param entityList      实体集
     * @param targetField     注解目标属性
     * @throws PoiOoxmlPlusException 异常
     */
    private void titleRender(Annotation annotation,
                             WorkbookCommand workbookCommand,
                             Class<?> clazz,
                             List<?> entityList,
                             Field targetField) throws PoiOoxmlPlusException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        // 判断注解是否存在处理器
        if (annotationClass.isAnnotationPresent(Processor.class)) {
            Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);

            // 处理器必须实现WorkbookProcessor
            Class<?> processor = processorAnnotation.value();
            if (WorkbookProcessor.class.isAssignableFrom(processor)) {
                try {
                    WorkbookProcessor workbookProcessor = (WorkbookProcessor) processor.newInstance();
                    workbookProcessor.titleRender(annotationEntity, workbookCommand, entityList, clazz, targetField);
                } catch (IllegalAccessException | InstantiationException e) {
                    throw new PoiOoxmlPlusException("注解处理器实例化失败;".concat(e.getMessage()));
                }
            }
        }
    }

    /**
     * 解析数据处理器
     *
     * @param annotation      注解
     * @param workbookCommand 基础命令
     * @param object          实体
     * @param entityList      实体集
     * @param targetField     注解目标属性
     * @param value           注解目标属性值
     * @return 注解处理器更新后的值
     * @throws PoiOoxmlPlusException 异常
     */
    private Object bodyRender(Annotation annotation,
                              WorkbookCommand workbookCommand,
                              Class<?> clazz,
                              Object object,
                              List<?> entityList,
                              Field targetField,
                              Object value) throws PoiOoxmlPlusException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        // 判断注解是否存在处理器
        if (annotationClass.isAnnotationPresent(Processor.class)) {
            Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);

            // 处理器必须实现WorkbookProcessor
            Class<?> processor = processorAnnotation.value();
            if (WorkbookProcessor.class.isAssignableFrom(processor)) {
                try {
                    WorkbookProcessor workbookProcessor = (WorkbookProcessor) processor.newInstance();
                    value = workbookProcessor.bodyRender(annotationEntity, workbookCommand, object, entityList, clazz, targetField, value);
                } catch (IllegalAccessException | InstantiationException e) {
                    throw new PoiOoxmlPlusException("注解处理器实例化失败;".concat(e.getMessage()));
                }
            }
        }
        return value;
    }

    /**
     * Poi属性缓存
     */
    private static final Map<Class<?>, List<Field>> poiColumnFieldListCache = new HashMap<>(16);

    /**
     * 获取有@PoiColumn注解的属性集合<BR/>
     * 解析PoiColumn的sort，进行排序
     *
     * @param clazz 类
     * @return 获取有@PoiColumn注解的属性集合
     */
    private List<Field> poiColumnFieldList(Class<?> clazz) throws PoiOoxmlPlusException {
        // 静态缓存
        List<Field> poiColumnFieldList = poiColumnFieldListCache.get(clazz);
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
        poiColumnFieldListCache.put(clazz, poiColumnFieldList);
        return poiColumnFieldList;
    }
}