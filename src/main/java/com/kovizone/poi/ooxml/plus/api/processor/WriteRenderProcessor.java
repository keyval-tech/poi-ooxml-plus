package com.kovizone.poi.ooxml.plus.api.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

/**
 * 数据表格输出处理器接口
 *
 * @author 11061
 */
public interface WriteRenderProcessor<A extends Annotation> extends BaseProcessor<A> {

    /**
     * 表头列渲染
     *
     * @param annotation   注解类实体
     * @param excelCommand 基础命令
     * @param clazz        实体类
     */
    default void headerRender(A annotation,
                              ExcelCommand excelCommand,
                              Class<?> clazz) {
    }

    /**
     * 数据标题渲染
     *
     * @param annotation   注解类实体
     * @param excelCommand 基础命令
     * @param targetField  注解目标属性
     */
    default void dataTitleRender(A annotation,
                                 ExcelCommand excelCommand,
                                 Field targetField) {
    }

    /**
     * 数据体渲染
     *
     * @param annotation   注解类实体
     * @param excelCommand 基础命令
     * @param targetField  注解目标属性
     * @param columnValue  注解目标属性值
     * @return 处理后的数据单元格值
     */
    default Object dataBodyRender(A annotation,
                                  ExcelCommand excelCommand,
                                  Field targetField,
                                  Object columnValue) {
        return columnValue;
    }

}
