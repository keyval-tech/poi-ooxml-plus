package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteSuffixProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定后缀属性
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteSuffixProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteSuffix {

    /**
     * 输出值拼接后缀<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     *
     * @return 拼接后缀集
     */
    String value();

}
