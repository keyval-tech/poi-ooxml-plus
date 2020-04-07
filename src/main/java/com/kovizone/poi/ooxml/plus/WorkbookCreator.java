package com.kovizone.poi.ooxml.plus;

import com.kovizone.poi.ooxml.plus.anno.ColumnConfig;
import com.kovizone.poi.ooxml.plus.anno.HeaderRender;
import com.kovizone.poi.ooxml.plus.anno.Processor;
import com.kovizone.poi.ooxml.plus.anno.SheetConfig;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import com.kovizone.poi.ooxml.plus.processor.ColumnValueProcessors;
import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;
import com.kovizone.poi.ooxml.plus.style.WorkbookStyle;
import com.kovizone.poi.ooxml.plus.style.impl.WorkbookDefaultStyle;
import com.kovizone.poi.ooxml.plus.util.StringUtils;
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

    /**
     * 工作表样式
     */
    private WorkbookStyle workbookStyle;
    /**
     * 表头标题样式
     */
    private CellStyle headerTitleCellStyle;
    /**
     * 表头标题行样式
     */
    private CellStyle headerTitleRowStyle;
    /**
     * 表头副标题样式
     */
    private CellStyle headerSubtitleCellStyle;
    /**
     * 表头副标题行样式
     */
    private CellStyle headerSubtitleRowStyle;
    /**
     * 数据标题样式
     */
    private CellStyle dateTitleCellStyle;
    /**
     * 数据标题行样式
     */
    private CellStyle dateTitleRowStyle;
    /**
     * 数据内容样式
     */
    private CellStyle dateBodyCellStyle;
    /**
     * 数据内容行样式
     */
    private CellStyle dateBodyRowStyle;
    /**
     * 工作表命令
     */
    private WorkbookCommand workbookCommand;

    /**
     * Poi属性缓存
     */
    private static final Map<Class<?>, List<Field>> poiColumnFieldListCache = new HashMap<>(16);

    /**
     * 实体类构造器<BR/>
     * 注入默认样式
     */
    public WorkbookCreator() {
        super();
        this.workbookStyle = new WorkbookDefaultStyle();
    }

    /**
     * 实体类构造器<BR/>
     * 注入自定义样式
     */
    public WorkbookCreator(WorkbookStyle workbookStyle) {
        super();
        this.workbookStyle = workbookStyle;
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
        List<Field> poiColumnFieldList = poiColumnFieldList(clazz);
        workbookCommand = new WorkbookCommand(workbook, poiColumnFieldList.size() - 1, textReplaceMap);
        int dateIndex = 0;

        sheetCycle:
        while (true) {

            Annotation[] clazzAnnotations = clazz.getDeclaredAnnotations();
            for (Annotation clazzAnnotation : clazzAnnotations) {
                Class<? extends Annotation> annotation = clazzAnnotation.annotationType();
                // 判断注解是否存在处理器
                if (annotation.isAnnotationPresent(Processor.class)) {
                    Processor processorAnnotation = annotation.getDeclaredAnnotation(Processor.class);
                    Object annotationEntity = clazz.getDeclaredAnnotation(clazzAnnotation.annotationType());

                    // 处理器必须实现WorkbookProcessor
                    Class<?> processor = processorAnnotation.value();
                    if (WorkbookProcessor.class.isAssignableFrom(processor)) {
                        try {
                            WorkbookProcessor workbookProcessor = (WorkbookProcessor) processor.newInstance();
                            workbookProcessor.render(annotationEntity, workbookCommand, clazz, entityList, null, null);
                        } catch (IllegalAccessException | InstantiationException e) {
                            throw new PoiOoxmlPlusException("注解处理器实例化失败;".concat(e.getMessage()));
                        }
                    }
                }
            }

            // 数据标题
            workbookCommand.createRow(dateTitleRowStyle);
            for (Field field : poiColumnFieldList) {
                Cell cell = workbookCommand.createCell();
                if (dateTitleCellStyle != null) {
                    cell.setCellStyle(dateTitleCellStyle);
                }
                renderDateTitle(cell, field);
            }

            // 数据主体
            for (; dateIndex < entityList.size(); dateIndex++) {
                Object entity = entityList.get(dateIndex);

                if (workbookCommand.currentRowIndex() >= maxLength) {
                    // 达到最大行数，新增工作簿
                    continue sheetCycle;
                }

                workbookCommand.createRow(dateBodyRowStyle);

                for (Field field : poiColumnFieldList) {
                    Object value = null;
                    try {
                        value = field.get(entity);
                    } catch (IllegalAccessException e) {
                        throw new PoiOoxmlPlusException("读取值失败");
                    }
                    Annotation[] fieldAnnotations = field.getDeclaredAnnotations();
                    processor(fieldAnnotations, workbookCommand, clazz, entityList, field, value);


                    Cell cell = workbookCommand.createCell();
                    if (dateBodyCellStyle != null) {
                        cell.setCellStyle(dateBodyCellStyle);
                    }
                    ColumnValueProcessors.setCellValue(cell, entity, field);
                }
            }
            workbookCommand.lateRender();
            break;
        }
    }

    private void processor(Annotation[] annotations,
                           WorkbookCommand workbookCommand,
                           Class<?> clazz,
                           List<?> entityList,
                           Field targetField,
                           Object value) throws PoiOoxmlPlusException {
        for (Annotation annotation : annotations) {
            Class<? extends Annotation> annotationClass = annotation.annotationType();
            // 判断注解是否存在处理器
            if (annotationClass.isAnnotationPresent(Processor.class)) {
                Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
                Object annotationEntity = clazz.getDeclaredAnnotation(annotation.annotationType());

                // 处理器必须实现WorkbookProcessor
                Class<?> processor = processorAnnotation.value();
                if (WorkbookProcessor.class.isAssignableFrom(processor)) {
                    try {
                        WorkbookProcessor workbookProcessor = (WorkbookProcessor) processor.newInstance();
                        workbookProcessor.render(annotationEntity, workbookCommand, clazz, entityList, targetField, value);
                    } catch (IllegalAccessException | InstantiationException e) {
                        throw new PoiOoxmlPlusException("注解处理器实例化失败;".concat(e.getMessage()));
                    }
                }
            }
        }

    }

    /**
     * 渲染数据标题
     *
     * @param cell  列
     * @param field 属性
     */
    private void renderDateTitle(Cell cell, Field field) throws PoiOoxmlPlusException {
        ColumnConfig columnConfig = field.getDeclaredAnnotation(ColumnConfig.class);
        ColumnValueProcessors.setCellValue(cell, field, columnConfig.title());
        int cellWidth = columnConfig.width();
        if (cellWidth != 0) {
            workbookCommand.setCurrentCellWidth(cellWidth);
        }
    }

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