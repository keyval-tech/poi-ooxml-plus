package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataBodyProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteDataTitleProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteHeaderProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteSheetInitProcessor;
import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import com.kovizone.poi.ooxml.plus.util.ReflexUtils;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 输出数据体处理器工厂
 *
 * @author KoviChen
 */
public class ProcessorFactory {

    /**
     * 解析表头处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @throws ExcelWriteException 异常
     */
    @SuppressWarnings("unchecked")
    public static void sheetInitProcessor(Annotation annotation,
                                          ExcelCommand excelCommand,
                                          Class<?> clazz) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteSheetInitProcessor writeSheetInitProcessor = getProcessor(annotationClass, WriteSheetInitProcessor.class);
        if (writeSheetInitProcessor != null) {
            Annotation annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeSheetInitProcessor.sheetInitProcess(annotationEntity,
                    excelCommand,
                    clazz);
        }
    }

    /**
     * 解析表头处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @throws ExcelWriteException 异常
     */
    public static void headerProcessor(Annotation annotation,
                                       ExcelCommand excelCommand,
                                       Class<?> clazz) throws ExcelWriteException {

        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteHeaderProcessor writeHeaderProcessor = getProcessor(annotationClass, WriteHeaderProcessor.class);
        if (writeHeaderProcessor != null) {
            excelCommand.createRow();
            Object annotationEntity = clazz.getDeclaredAnnotation(annotationClass);
            writeHeaderProcessor.headerProcess((Annotation) annotationEntity,
                    excelCommand,
                    clazz);
        }
    }

    /**
     * 解析数据标题处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param targetField  注解目标属性
     * @throws ExcelWriteException 异常
     */
    public static void dateTitleProcessor(Annotation annotation,
                                          ExcelCommand excelCommand,
                                          Field targetField) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDataTitleProcessor writeDataTitleProcessor = getProcessor(annotationClass, WriteDataTitleProcessor.class);
        if (writeDataTitleProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            writeDataTitleProcessor.dataTitleProcess((Annotation) annotationEntity,
                    excelCommand,
                    targetField);
        }
    }

    /**
     * 解析数据处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param targetField  注解目标属性
     * @param columnValue  注解目标属性值
     * @return 注解处理器更新后的值
     * @throws ExcelWriteException 异常
     */
    public static Object dateBodyProcessor(Annotation annotation,
                                           ExcelCommand excelCommand,
                                           Field targetField,
                                           Object columnValue) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteDataBodyProcessor writeDataBodyProcessor = getProcessor(annotationClass, WriteDataBodyProcessor.class);
        if (writeDataBodyProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            if (annotationEntity == null) {
                annotationEntity = excelCommand.getEntityList().get(excelCommand.currentEntityListIndex()).getClass().getDeclaredAnnotation(annotationClass);
            }
            columnValue = writeDataBodyProcessor.dataBodyProcess((Annotation) annotationEntity,
                    excelCommand,
                    targetField,
                    columnValue);
        }
        return columnValue;
    }

    private final static List<Class<?>> NOT_PROCESSOR_CACHE = new ArrayList<>();

    private final static Map<String, Object> PROCESSOR_CACHE = new HashMap<>();

    private final static String NOT_FOUND_PROCESSOR_PLACEHOLDER = "not found processor";

    private static <P> P getProcessor(Class<? extends Annotation> annotationClass, Class<P> processorClass) throws ExcelWriteException {
        if (NOT_PROCESSOR_CACHE.contains(annotationClass)) {
            return null;
        }
        String cacheKey = String.format("%s_%s", annotationClass.getSimpleName(), processorClass.getSimpleName());
        Object processorObject = PROCESSOR_CACHE.get(cacheKey);
        if (processorObject == null) {
            // 判断注解是否存在处理器
            if (annotationClass.isAnnotationPresent(Processor.class)) {
                Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
                Class<?> processor = processorAnnotation.value();
                if (processorClass.isAssignableFrom(processor)) {
                    try {
                        processorObject = ReflexUtils.newInstance(processor);
                        PROCESSOR_CACHE.put(cacheKey, processorObject);
                    } catch (ReflexException e) {
                        e.printStackTrace();
                        throw new ExcelWriteException("构造处理器失败;" + e.getMessage());
                    }
                } else {
                    PROCESSOR_CACHE.put(cacheKey, NOT_FOUND_PROCESSOR_PLACEHOLDER);
                }
            } else {
                NOT_PROCESSOR_CACHE.add(annotationClass);
            }
        } else if (NOT_FOUND_PROCESSOR_PLACEHOLDER.equals(String.valueOf(processorObject))) {
            return null;
        }
        return (P) processorObject;
    }
}
