package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.WorkbookCommand;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;

import java.lang.reflect.Field;
import java.util.List;

/**
 * 处理器接口
 *
 * @author 11061
 */
public interface WorkbookProcessor {

    /**
     * 表头渲染
     *
     * @param annotation      注解类实体
     * @param workbookCommand 基础命令
     * @param entityList      渲染实体集合
     * @param clazz           实体类
     * @throws PoiOoxmlPlusException 异常
     */
    default void headerRender(Object annotation,
                              WorkbookCommand workbookCommand,
                              List<?> entityList,
                              Class<?> clazz
    ) throws PoiOoxmlPlusException {
    }

    /**
     * 数据标题渲染
     *
     * @param annotation      注解类实体
     * @param workbookCommand 基础命令
     * @param entityList      渲染实体集合
     * @param clazz           实体类
     * @param targetField     注解目标属性
     * @throws PoiOoxmlPlusException 异常
     */
    default void titleRender(Object annotation,
                             WorkbookCommand workbookCommand,
                             List<?> entityList,
                             Class<?> clazz,
                             Field targetField
    ) throws PoiOoxmlPlusException {
    }

    /**
     * 数据体渲染
     *
     * @param annotation       注解类实体
     * @param workbookCommand  基础命令
     * @param object           实体
     * @param entityList       渲染实体集合
     * @param clazz            实体类
     * @param targetField      注解目标属性
     * @param targetFieldValue 注解目标属性值
     * @throws PoiOoxmlPlusException 异常
     */
    default Object bodyRender(Object annotation,
                              WorkbookCommand workbookCommand,
                              Object object,
                              List<?> entityList,
                              Class<?> clazz,
                              Field targetField,
                              Object targetFieldValue
    ) throws PoiOoxmlPlusException {
        return targetFieldValue;
    }

}
