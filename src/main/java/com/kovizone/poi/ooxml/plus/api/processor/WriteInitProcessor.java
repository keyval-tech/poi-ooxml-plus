package com.kovizone.poi.ooxml.plus.api.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelInitCommand;

import java.lang.annotation.Annotation;

/**
 * 数据表格输出处理器接口
 *
 * @author 11061
 */
public interface WriteInitProcessor<A extends Annotation> extends BaseProcessor<A> {

    /**
     * Sheet初始化渲染
     *
     * @param annotation       注解类实体
     * @param excelInitCommand 初始化命令
     * @param clazz            实体类
     */
    default void init(A annotation,
                      ExcelInitCommand excelInitCommand,
                      Class<?> clazz) {
    }
}
