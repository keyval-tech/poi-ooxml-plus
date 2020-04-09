package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.command.ExcelCommand;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 数据表格处理器接口
 *
 * @author 11061
 */
public interface WriteDateBodyProcessor {


    /**
     * 处理
     *
     * @param annotation      注解类实体
     * @param excelCommand 基础命令
     * @param entityListIndex 渲染实体集合当前索引
     * @param entityList      渲染实体集合
     * @param targetField     注解目标属性
     * @param columnValue     注解目标属性值
     * @return 处理后的值
     * @throws PoiOoxmlPlusException 异常
     */
    Object dateBodyProcess(Object annotation,
                           ExcelCommand excelCommand,
                           List<?> entityList,
                           int entityListIndex,
                           Field targetField,
                           Object columnValue
    ) throws PoiOoxmlPlusException;

}
