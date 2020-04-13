package com.kovizone.poi.ooxml.plus.api.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 数据标题处理器接口
 *
 * @author KoviChen
 */
public interface WriteDataTitleProcessor {

    /**
     * 处理
     *
     * @param annotation      注解类实体
     * @param excelCommand 基础命令
     * @param entityList      渲染实体集合
     * @param targetField     注解目标属性
     */
    void dataTitleProcess(Object annotation,
                          ExcelCommand excelCommand,
                          List<?> entityList,
                          Field targetField
    );
}
