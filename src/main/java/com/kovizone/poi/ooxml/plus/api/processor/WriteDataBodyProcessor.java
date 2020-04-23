package com.kovizone.poi.ooxml.plus.api.processor;


import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

/**
 * 数据表格处理器接口
 *
 * @author 11061
 */
public interface WriteDataBodyProcessor<A extends Annotation> {


    /**
     * 处理
     *
     * @param annotation      注解类实体
     * @param excelCommand    基础命令
     * @param targetField     注解目标属性
     * @param columnValue     注解目标属性值
     * @return 处理后的值
     */
    Object dataBodyProcess(A annotation,
                           ExcelCommand excelCommand,
                           Field targetField,
                           Object columnValue
    );

}
