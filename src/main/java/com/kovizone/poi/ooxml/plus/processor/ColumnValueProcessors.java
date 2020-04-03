package com.kovizone.poi.ooxml.plus.processor;

import com.kovizone.poi.ooxml.plus.anno.*;
import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 单元格插值处理器
 *
 * @author KoviChen
 */
public class ColumnValueProcessors {

    /**
     * 设置单元格值（工厂）<BR/>
     *
     * @param cell   行
     * @param entity 实体
     * @param field  属性
     */
    public static void setCellValue(Cell cell, Object entity, Field field) throws PoiOoxmlPlusException {
        setCellValue(cell, entity, field, null);
    }

    /**
     * 设置单元格值（工厂）<BR/>
     *
     * @param cell  行
     * @param field 属性
     * @param value 值
     */
    public static void setCellValue(Cell cell, Field field, Object value) throws PoiOoxmlPlusException {
        setCellValue(cell, null, field, value);
    }

    /**
     * 设置单元格值（工厂）<BR/>
     *
     * @param cell   行
     * @param entity 实体，允许为空
     * @param field  属性
     * @param value  值
     */
    private static void setCellValue(Cell cell, Object entity, Field field, Object value) throws PoiOoxmlPlusException {
        if (value == null) {
            // 显示条件，不满足显示条件时，value = ""
            if (field.isAnnotationPresent(Criteria.class)) {
                Criteria criteria = field.getDeclaredAnnotation(Criteria.class);
                for (String condition : criteria.value()) {
                    value = ConditionProcessors.processorFactory(entity, condition);
                }
            }
            if (value == null) {
                try {
                    value = field.get(entity);
                } catch (IllegalAccessException e) {
                    throw new PoiOoxmlPlusException("读取属性值失败");
                }
            }
            if (field.isAnnotationPresent(Substitute.class)) {
                Substitute substitute = field.getDeclaredAnnotation(Substitute.class);
                for (String condition : substitute.condition()) {
                    value = SubstituteProcessors.processorFactory(entity, value, condition, substitute.value());
                }
            }
        }
        if (value == null) {
            if (field.isAnnotationPresent(DefaultValue.class)) {
                value = field.getDeclaredAnnotation(DefaultValue.class).value();
            }
        }

        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
            return;
        }
        if (value instanceof Short
                || value instanceof Integer
                || value instanceof Long
                || value instanceof Float
                || value instanceof Double) {
            setCellNumberValue(cell, entity, field, value);
            return;
        }
        if (value instanceof Date) {
            setCellDateValue(cell, entity, field, value);
            return;
        }
        setCellStringValue(cell, entity, field, value);
    }

    /**
     * 设置日期类单元格值<BR/>
     * 配合PoiStringReplace转移注解
     *
     * @param cell   列
     * @param entity 实体，允许为空
     * @param field  属性
     * @param value  值
     */
    private static void setCellDateValue(Cell cell, Object entity, Field field, Object value) throws PoiOoxmlPlusException {
        if (field.isAnnotationPresent(DateFormat.class)) {
            DateFormat dateFormat = field.getDeclaredAnnotation(DateFormat.class);
            String str = new SimpleDateFormat(dateFormat.value()).format(((Date) value));
            setCellStringValue(cell, entity, field, str);
            return;
        }
        cell.setCellValue((Date) value);
    }

    /**
     * 设置文本类单元格值<BR/>
     * 配合PoiStringReplace转移注解
     *
     * @param cell   列
     * @param entity 实体，允许为空
     * @param field  属性
     * @param value  值
     */
    private static void setCellStringValue(Cell cell, Object entity, Field field, Object value) throws PoiOoxmlPlusException {
        String valueString = String.valueOf(value);
        if (field.isAnnotationPresent(StringReplace.class)) {
            StringReplace stringReplace = field.getDeclaredAnnotation(StringReplace.class);
            String[] targets = stringReplace.target();
            String[] replacements = stringReplace.replacement();
            if (targets.length != replacements.length) {
                throw new PoiOoxmlPlusException(field.getName() + "的@PoiStringReplace配置有误，target和replacement长度不一致");
            }
            for (int i = 0; i < targets.length; i++) {
                String target = targets[i];
                if (valueString.contains(target)) {
                    String replacement = replacements[i];
                    valueString = valueString.replace(target, replacement);
                }
            }
        }

        // 拼接前缀后缀
        if (value != null) {
            if (entity == null) {
                cell.setCellValue(valueString);
            } else {
                cell.setCellValue(ConcatProcessors.concat(entity, field, valueString));
            }
        }
    }

    /**
     * 设置数字类单元格值<BR/>
     * 配合PoiNumberReplace转移注解
     *
     * @param cell   列
     * @param entity 实体，允许为空
     * @param field  属性
     * @param value  值
     */
    private static void setCellNumberValue(Cell cell, Object entity, Field field, Object value) throws PoiOoxmlPlusException {
        if (field.isAnnotationPresent(NumberMapper.class)) {
            NumberMapper numberMapper = field.getDeclaredAnnotation(NumberMapper.class);
            int[] target = numberMapper.target();
            for (int i = 0; i < target.length; i++) {
                int t = target[i];
                if ((int) value == t) {
                    String valueString = numberMapper.replacement()[i];
                    setCellStringValue(cell, entity, field, valueString);
                    return;
                }
            }
        }
        if (value instanceof Short) {
            cell.setCellValue((Short) value);
            return;
        }
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
            return;
        }
        if (value instanceof Long) {
            cell.setCellValue((Long) value);
            return;
        }
        if (value instanceof Float) {
            cell.setCellValue((Float) value);
            return;
        }
        if (value instanceof Double) {
            cell.setCellValue((Double) value);
            return;
        }
        setCellStringValue(cell, entity, field, String.valueOf(value));
    }
}
