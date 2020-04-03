package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.processor.impl.HeaderRenderProcessors;

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
@Processor(HeaderRenderProcessors.class)
public @interface HeaderRender {

    /**
     * 表头标题
     *
     * @return 标题
     */
    String headerTitle() default "";

    /**
     * 表头副标题
     *
     * @return 副标题
     */
    String headerSubtitle() default "";

    /**
     * 标题单元格高度
     *
     * @return 标题单元格高度
     */
    short headerTitleHeight() default 1000;

    /**
     * 副标题单元格高度
     *
     * @return 副标题单元格高度
     */
    short headerSubtitleHeight() default 500;

}
