package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WritePrefixProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定前缀字段
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WritePrefixProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WritePrefix {

    /**
     * 输出值拼接前缀，
     * 使用#[FieldName]指定属性的值，
     *
     * @return 拼接前缀集
     */
    String value();

}
