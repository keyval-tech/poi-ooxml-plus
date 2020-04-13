package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.exception.ExpressionParseException;
import com.kovizone.poi.ooxml.plus.exception.ExcelWriteException;
import com.kovizone.poi.ooxml.plus.exception.ReflexException;
import org.springframework.expression.EvaluationContext;
import org.springframework.expression.Expression;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 表达式工具类
 *
 * @author KoviChen
 */
public class ElUtils {

    /**
     * 表达式解析器
     */
    private static ExpressionParser expressionParser = new SpelExpressionParser();

    /**
     * 数组获取对象操作符（假）
     */
    private static final String ARRAY_GET_FALSE = "[i]";
    /**
     * 数组获取对象操作符（真）
     */
    private static final String ARRAY_GET_TRUE = "[#i]";

    /**
     * 集合获取对象方法（假）
     */
    private static final String LIST_METHOD_GET_FALSE = ".get(i)";

    /**
     * 集合获取对象方法（真）
     */
    private static final String LIST_METHOD_GET_TRUE = ".get(#i)";

    /**
     * 实体集占位符
     */
    private static final String LIST_PLACEHOLDER = "#list";

    /**
     * 当前实体占位符
     */
    private static final String THIS_PLACEHOLDER = "#this";

    /**
     * 参数名：实体集
     */
    private static final String PARAM_LIST = "list";

    /**
     * 参数名：实体集当前索引
     */
    private static final String PARAM_INDEX = "i";

    /**
     * 转义Map
     */
    private static final Map<String, String> ESCAPE_MAP = new HashMap<>();

    static {
        ESCAPE_MAP.put(
                // #this
                THIS_PLACEHOLDER,
                // #list[#i]
                LIST_PLACEHOLDER.concat(ARRAY_GET_TRUE)
        );

        ESCAPE_MAP.put(
                // #list[i]
                LIST_PLACEHOLDER.concat(ARRAY_GET_FALSE),
                // #list[#i]
                LIST_PLACEHOLDER.concat(ARRAY_GET_TRUE)
        );

        ESCAPE_MAP.put(
                // #list.get(i)
                LIST_PLACEHOLDER.concat(LIST_METHOD_GET_FALSE),
                // #list.get(#i)
                LIST_PLACEHOLDER.concat(LIST_METHOD_GET_TRUE)
        );
    }

    /**
     * 解析表达式转布偶值
     *
     * @param expressionString 表达式字符串
     * @param entityList       实体集
     * @param index            实体集当前索引
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    public static boolean parseBoolean(String expressionString, List<?> entityList, int index) throws ExpressionParseException {
        try {
            Boolean result = parse(expressionString, entityList, index, Boolean.class);
            if (result != null) {
                return result;
            }
            throw new ExcelWriteException("解析EL表达式失败：" + expressionString);
        } catch (Exception e) {
            throw new ExpressionParseException("解析EL表达式失败：" + expressionString + ";" + e.getMessage());
        }
    }

    /**
     * 解析表达式
     *
     * @param expressionString 表达式字符串
     * @param entityList       实体集
     * @param index            实体集当前索引
     * @param clazz            解析结果类
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    public static <T> T parse(String expressionString, List<?> entityList, int index, Class<T> clazz) throws ExpressionParseException {
        try {
            Expression expression = expressionParser.parseExpression(escape(expressionString));
            return expression.getValue(paramMap(entityList, index), clazz);
        } catch (Exception e) {
            if (clazz.equals(String.class) && !expressionString.contains("+") && !expressionString.contains(".concat")) {
                Expression expression = expressionParser.parseExpression(escape("'" + expressionString + "'"));
                return expression.getValue(paramMap(entityList, index), clazz);
            }
            throw new ExpressionParseException("解析EL表达式失败：" + expressionString + ";" + e.getMessage());
        }
    }

    /**
     * 解析表达式转字符串
     *
     * @param expressionString 表达式字符串
     * @param entityList       实体集
     * @param index            实体集当前索引
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    public static String parseString(String expressionString, List<?> entityList, int index) throws ExpressionParseException {
        return parse(expressionString, entityList, index, String.class);
    }

    /**
     * 转义
     *
     * @param arg 原文
     * @return 转义问
     */
    private static String escape(String arg) {
        String str = arg;
        Set<Map.Entry<String, String>> entrySet = ESCAPE_MAP.entrySet();
        for (Map.Entry<String, String> entry : entrySet) {
            if (str.contains(entry.getKey())) {
                str = str.replace(entry.getKey(), entry.getValue());
            }
        }
        return str;
    }

    private static EvaluationContext paramMap(List<?> entityList, int index) throws ExpressionParseException {
        EvaluationContext context = new StandardEvaluationContext();
        Object entity = entityList.get(index);
        Field[] fields = entity.getClass().getDeclaredFields();
        for (Field field : fields) {
            try {
                context.setVariable(field.getName(), ReflexUtils.getValue(entity, field));
            } catch (ReflexException e) {
                throw new ExpressionParseException("读取当前实体值失败");
            }
        }
        context.setVariable(PARAM_LIST, entityList);
        context.setVariable(PARAM_INDEX, index);

        return context;
    }
}
