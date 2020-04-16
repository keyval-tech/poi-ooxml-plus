package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteSubstituteProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定在特殊条件下的替补字段
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteSubstituteProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteSubstitute {

    /**
     * 替补条件<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     *
     * @return 替补条件
     */
    String criteria();

    /**
     * 替补值<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     * 否则为常量
     *
     * @return 替补属性名
     */
    String value();
}
