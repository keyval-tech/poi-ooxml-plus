package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteHeaderProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定表头数据
 *
 * @author KoviChen
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Processor(WriteHeaderProcessors.class)
public @interface WriteHeader {

    /**
     * 表头标题
     *
     * @return 标题
     */
    String value() default "";

    /**
     * 标题单元格高度
     *
     * @return 标题单元格高度
     */
    short height() default 1000;

}
