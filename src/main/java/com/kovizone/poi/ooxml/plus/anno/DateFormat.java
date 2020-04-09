package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.anno.base.Processor;
import com.kovizone.poi.ooxml.plus.processor.impl.DateFormatProcessors;

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
@Processor(DateFormatProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface DateFormat {

    String value() default "";

}
