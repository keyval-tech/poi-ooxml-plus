package com.kovizone.poi.ooxml.plus.api.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;

import java.util.List;

/**
 * 工作簿初始化处理器接口
 *
 * @author 11061
 */
public interface WriteSheetInitProcessor {
    /**
     * 处理
     *
     * @param annotation   注解类实体
     * @param excelCommand 基础命令
     * @param entityList   渲染实体集合
     * @param clazz        实体类
     */
    void sheetInitProcess(Object annotation,
                          ExcelCommand excelCommand,
                          List<?> entityList,
                          Class<?> clazz
    );

}
