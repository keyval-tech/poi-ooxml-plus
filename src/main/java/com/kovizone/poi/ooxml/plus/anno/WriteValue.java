package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteValueProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 输出到单元格的内容
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteValueProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteValue {

    /**
     * 显示条件表达式，
     * 使用#FieldName指定属性的值，
     *
     * @return 显示条件表达式
     */
    String[] value();

}
