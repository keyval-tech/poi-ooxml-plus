package com.kovizone.poi.ooxml.plus.api.converter;

/**
 * 数字转换器
 *
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
@FunctionalInterface
public interface NumericCellConverter<T> extends CellConverter<Double, T> {
}
