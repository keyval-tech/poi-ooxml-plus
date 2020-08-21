package com.kovizone.poi.ooxml.plus.api.converter;

/**
 * 转换器
 *
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
@FunctionalInterface
public interface CellConverter<S, T> {
    /**
     * {@code S}类型转为{@code T}类型
     *
     * @param var1 原类型
     * @return 转换后类型
     */
    T convert(S var1);
}
