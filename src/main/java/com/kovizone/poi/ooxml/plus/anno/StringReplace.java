package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.processor.StringReplaceProcessors;

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
@Processor(StringReplaceProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface StringReplace {

    /**
     * 关键字
     *
     * @return 关键字
     */
    String[] target();

    /**
     * 转义词
     *
     * @return 转义词
     */
    String[] replacement();

}
