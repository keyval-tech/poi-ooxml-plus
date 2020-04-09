package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.exception.PoiOoxmlPlusException;
import org.apache.poi.ss.formula.functions.T;
import org.mvel2.MVEL;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

/**
 * @author KoviChen
 */
public class MvelUtils {

    private static String escape(String expression) {
        if (expression.contains("is")) {
            expression = expression.replace(" is ", " == ");
        }
        if (expression.contains("IS")) {
            expression = expression.replace(" IS ", " == ");
        }
        return expression;
    }

    public static boolean parseBoolean(String expression, List<?> entityList, int index) throws PoiOoxmlPlusException {
        Object result = MVEL.eval(escape(expression), paramMap(entityList, index));
        if (!(result instanceof Boolean)) {
            throw new PoiOoxmlPlusException("表达式解析失败：" + expression);
        }
        return (Boolean) result;
    }

    public static String parseString(String expression, List<?> entityList, int index) throws PoiOoxmlPlusException {
        return MVEL.evalToString(escape(expression), paramMap(entityList, index));
    }

    public static Map<String, Object> paramMap(List<?> entityList, int index) throws PoiOoxmlPlusException {
        Map<String, Object> paramMap = new HashMap<>(16);

        paramMap.put("list", entityList);
        paramMap.put("i", index);

        Object entity = entityList.get(index);
        Field[] fields = entity.getClass().getDeclaredFields();
        for (Field field : fields) {
            paramMap.put(field.getName(), ReflexUtils.getValue(entity, field));
        }
        return paramMap;
    }
}
