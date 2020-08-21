package com.kovizone.poi.ooxml.plus.api.anno;

import com.kovizone.poi.ooxml.plus.api.converter.CellConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ReaderConverter {

    Class<? extends CellConverter<?, ?>>[] value();

}
