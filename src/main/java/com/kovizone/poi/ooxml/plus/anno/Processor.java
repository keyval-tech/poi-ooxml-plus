package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.processor.WorkbookProcessor;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定处理器
 *
 * @author KoviChen
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface Processor {

    Class<?> value();

}
