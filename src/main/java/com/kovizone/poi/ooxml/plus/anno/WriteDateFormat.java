package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteDataFormatProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 插入Date类型到单元格时时格式化为字符串
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteDataFormatProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteDateFormat {

    String value() default "";

}
