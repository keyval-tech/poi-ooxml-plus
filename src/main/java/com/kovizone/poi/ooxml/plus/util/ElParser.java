package com.kovizone.poi.ooxml.plus.util;

import com.kovizone.poi.ooxml.plus.exception.ExpressionParseException;
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
public class ElParser {

    /**
     * 表达式解析器
     */
    private static ExpressionParser expressionParser = new SpelExpressionParser();

    /**
     * 转义Map
     */
    private static final Map<String, String> ESCAPE_MAP = new HashMap<>();

    static {
        ESCAPE_MAP.put("#this", "#list[#i]");
        ESCAPE_MAP.put("#list[i]", "#list[#i]");
        ESCAPE_MAP.put("#list.get(i)", "#list.get(#i)");
    }

    /**
     * 解析表达式转布偶值
     *
     * @param expressionString 表达式字符串
     * @param index            实体集当前索引
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    public boolean parseBoolean(String expressionString, int index) throws ExpressionParseException {
        try {
            Boolean result = parse(expressionString, index, Boolean.class);
            if (result != null) {
                return result;
            }
            throw new ExpressionParseException("解析EL表达式失败：" + expressionString);
        } catch (Exception e) {
            throw new ExpressionParseException("解析EL表达式失败：" + expressionString + ";" + e.getMessage());
        }
    }

    /**
     * 解析表达式转字符串
     *
     * @param expressionString 表达式字符串
     * @param index            实体集当前索引
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    public String parseString(String expressionString, int index) throws
            ExpressionParseException {
        return parse(expressionString, index, String.class);
    }

    /**
     * 转义
     *
     * @param arg 原文
     * @return 转义问
     */
    private String escape(String arg) {
        String str = arg;
        Set<Map.Entry<String, String>> entrySet = ESCAPE_MAP.entrySet();
        for (Map.Entry<String, String> entry : entrySet) {
            if (str.contains(entry.getKey())) {
                str = str.replace(entry.getKey(), entry.getValue());
            }
        }
        return str;
    }

    private List<?> entityList;

    private Map<Integer, Map<String, Object>> valueMapCache = new HashMap<>();

    public ElParser(List<?> entityList) {
        super();
        this.entityList = entityList;
    }

    private EvaluationContext evaluationContext(int index) throws ExpressionParseException {
        EvaluationContext context = new StandardEvaluationContext();
        Map<String, Object> valueMap = valueMapCache.get(index);
        if (valueMap == null) {
            valueMap = new HashMap<>(16);
            Object entity = entityList.get(index);
            Field[] fields = ReflexUtils.getDeclaredFields(entity.getClass());
            for (Field field : fields) {
                try {
                    valueMap.put(field.getName(), ReflexUtils.getValue(entity, field));
                } catch (ReflexException e) {
                    throw new ExpressionParseException("读取当前实体值失败");
                }
            }
            valueMap.put("list", entityList);
            valueMap.put("i", index);
            valueMapCache.put(index, valueMap);
        }
        valueMap.forEach(context::setVariable);
        return context;
    }

    /**
     * 解析表达式
     *
     * @param expressionString 表达式字符串
     * @param index            实体集当前索引
     * @param clazz            解析结果类
     * @return 解析结果
     * @throws ExpressionParseException 异常
     */
    @SuppressWarnings("unchecked")
    public <T> T parse(String expressionString, int index, Class<T> clazz) throws ExpressionParseException {
        final String plusKey = "+";
        final String splicerMethod = ".concat";
        final String poundKey = "#";
        try {
            if (clazz.equals(String.class)
                    && !expressionString.contains(plusKey)
                    && !expressionString.contains(splicerMethod)
                    && !expressionString.contains(poundKey)) {
                return (T) expressionString;
            } else {
                Expression expression = expressionParser.parseExpression(escape(expressionString));
                return expression.getValue(evaluationContext(index), clazz);
            }
        } catch (
                Exception e) {
            throw new ExpressionParseException("解析EL表达式失败：" + expressionString + ";" + e.getMessage());
        }
    }

}
