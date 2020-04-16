package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteSubheaderProcessors;

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
@Processor(WriteSubheaderProcessors.class)
public @interface WriteSubheader {

    /**
     * 表头副标题
     *
     * @return 副标题
     */
    String value() default "";

    /**
     * 副标题单元格高度
     *
     * @return 副标题单元格高度
     */
    short height() default 500;

}
