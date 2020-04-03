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
     * 渲染
     *
     * @param annotation      注解类实体
     * @param workbookCommand 基础命令
     * @param clazz           类
     * @param entityList      渲染实体集合
     * @param targetField     注解目标属性（可能为空）
     */
    void render(Object annotation, WorkbookCommand workbookCommand, Class<?> clazz, List<?> entityList, Field targetField) throws PoiOoxmlPlusException;

}
