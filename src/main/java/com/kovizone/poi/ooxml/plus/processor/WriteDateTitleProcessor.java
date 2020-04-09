package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 数据标题处理器接口
 *
 * @author KoviChen
 */
public interface WriteDateTitleProcessor {

    /**
     * 处理
     *
     * @param annotation      注解类实体
     * @param excelCommand 基础命令
     * @param entityList      渲染实体集合
     * @param targetField     注解目标属性
     * @throws PoiOoxmlPlusException 异常
     */
    void dateTitleProcess(Object annotation,
                          ExcelCommand excelCommand,
                          List<?> entityList,
                          Field targetField
    ) throws PoiOoxmlPlusException;
}
