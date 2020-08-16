package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.api.processor.BaseProcessor;
import com.kovizone.poi.ooxml.plus.api.processor.WriteProcessor;
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
     * @param clazz        实体类
     * @throws ExcelWriteException 异常
     */
    @SuppressWarnings("unchecked")
    public static void headerProcessor(Annotation annotation,
                                       ExcelCommand excelCommand,
                                       Class<?> clazz) throws ExcelWriteException {

        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteProcessor writeProcessor = getProcessor(annotationClass, WriteProcessor.class);
        Annotation annotationEntity;
        try {
            annotationEntity = ReflexUtils.getDeclaredAnnotation(clazz, annotationClass.getSimpleName());
        } catch (ReflexException e) {
            throw new ExcelWriteException(e);
        }
        writeProcessor.headerProcess(annotationEntity,
                excelCommand,
                clazz);
    }

    /**
     * 解析数据标题处理器
     *
     * @param annotation   注解
     * @param excelCommand 基础命令
     * @param targetField  注解目标属性
     * @throws ExcelWriteException 异常
     */
    @SuppressWarnings("unchecked")
    public static void dateTitleProcessor(Annotation annotation,
                                          ExcelCommand excelCommand,
                                          Field targetField) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteProcessor writeProcessor = getProcessor(annotationClass, WriteProcessor.class);
        if (writeProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            writeProcessor.dataTitleProcess((Annotation) annotationEntity,
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
    @SuppressWarnings("unchecked")
    public static Object dateBodyProcessor(Annotation annotation,
                                           ExcelCommand excelCommand,
                                           Field targetField,
                                           Object columnValue) throws ExcelWriteException {
        Class<? extends Annotation> annotationClass = annotation.annotationType();
        WriteProcessor writeProcessor = getProcessor(annotationClass, WriteProcessor.class);
        if (writeProcessor != null) {
            Object annotationEntity = targetField.getDeclaredAnnotation(annotationClass);
            if (annotationEntity == null) {
                annotationEntity = excelCommand.getEntityList().get(excelCommand.currentEntityListIndex()).getClass().getDeclaredAnnotation(annotationClass);
            }
            columnValue = writeProcessor.dataBodyProcess((Annotation) annotationEntity,
                    excelCommand,
                    targetField,
                    columnValue);
        }
        return columnValue;
    }

    private final static List<Class<?>> NOT_PROCESSOR_CACHE = new ArrayList<>();

    private final static Map<String, Object> PROCESSOR_CACHE = new HashMap<>();

    private final static String NOT_FOUND_PROCESSOR_PLACEHOLDER = "not found processor";

    @SuppressWarnings("unchecked")
    private static <P extends BaseProcessor<?>> P getProcessor(Class<? extends Annotation> annotationClass, Class<P> processorClass) throws ExcelWriteException {
        if (NOT_PROCESSOR_CACHE.contains(annotationClass)) {
            return null;
        }
        String cacheKey = String.format("%s_%s", annotationClass.getSimpleName(), processorClass.getSimpleName());
        Object processorObject = PROCESSOR_CACHE.get(cacheKey);
        if (processorObject == null) {
            // 判断注解是否存在处理器
            if (annotationClass.isAnnotationPresent(Processor.class)) {
                Processor processorAnnotation = annotationClass.getDeclaredAnnotation(Processor.class);
                Class<? extends BaseProcessor<?>>[] processors = processorAnnotation.value();
                for (Class<? extends BaseProcessor<?>> processor : processors) {
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
