package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import org.mvel2.MVEL;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author KoviChen
 */
public class MvelUtils {

    public static boolean parseBoolean(String expression, Object object) throws PoiOoxmlPlusException {
        Map<String, Object> paramMap = new HashMap<>();
        List<String> fieldNameList = FieldUtils.getFieldNameList(expression);
        for (String fieldName : fieldNameList) {
            try {
                Field field = object.getClass().getDeclaredField(fieldName);
                Object value = field.get(object);
                paramMap.put("#[" + fieldName + "]", value);
            } catch (NoSuchFieldException e) {
                throw new PoiOoxmlPlusException("没有找到属性：" + fieldName);
            } catch (IllegalAccessException e) {
                throw new PoiOoxmlPlusException("读取属性值失败：" + fieldName);
            }
        }
        Object result = MVEL.eval(expression, paramMap);
        if (!(result instanceof Boolean)) {
            throw new PoiOoxmlPlusException("表达式解析失败：" + expression);
        }
        return (Boolean) result;
    }

}
