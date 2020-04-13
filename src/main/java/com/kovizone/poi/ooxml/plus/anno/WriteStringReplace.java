package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.impl.WriteStringReplaceProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * String类型参数转义
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteStringReplaceProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteStringReplace {

    /**
     * 正则表达式
     *
     * @return 关键字
     */
    String[] regex() default "";

    /**
     * 转义词
     *
     * @return 转义词
     */
    String[] replacement();

}
