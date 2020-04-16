package com.kovizone.poi.ooxml.plus.api.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;

import java.lang.annotation.Annotation;
import java.util.List;

/**
 * 表头处理器接口
 *
 * @author KoviChen
 */
public interface WriteHeaderProcessor<A extends Annotation> {
    /**
     * 处理
     *
     * @param annotation      注解类实体
     * @param excelCommand 基础命令
     * @param entityList      渲染实体集合
     * @param clazz           实体类
     */
    void headerProcess(A annotation,
                       ExcelCommand excelCommand,
                       List<?> entityList,
                       Class<?> clazz
    );
}
