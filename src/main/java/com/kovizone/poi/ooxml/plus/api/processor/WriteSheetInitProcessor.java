package com.kovizone.poi.ooxml.plus.api.processor;


import com.kovizone.poi.ooxml.plus.command.ExcelCommand;

import java.lang.annotation.Annotation;

/**
 * 工作簿初始化处理器接口
 *
 * @author 11061
 */
public interface WriteSheetInitProcessor<A extends Annotation> {
    /**
     * 处理
     *
     * @param annotation   注解类实体
     * @param excelCommand 基础命令
     * @param clazz        实体类
     */
    void sheetInitProcess(A annotation,
                          ExcelCommand excelCommand,
                          Class<?> clazz
    );

}
