package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteStringReplaceAllProcessors;

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
@Processor(WriteStringReplaceAllProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteStringReplaceAll {

    /**
     * 关键字
     *
     * @return 关键字
     */
    String[] target() default "";

    /**
     * 转义词
     *
     * @return 转义词
     */
    String[] replacement();

}
