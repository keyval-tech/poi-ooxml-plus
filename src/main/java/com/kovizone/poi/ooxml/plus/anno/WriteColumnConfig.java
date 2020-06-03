package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteColumnConfigProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <p>指定为Excel列字段</p>
 *
 * <p>注解处理器{@link WriteColumnConfigProcessors}</p>
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(WriteColumnConfigProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface WriteColumnConfig {

    /**
     * 指定显示列名
     *
     * @return 列名
     */
    String value();

    /**
     * 排序，由小到大
     *
     * @return 排序值
     */
    int sort() default 0;

    /**
     * 列宽设置，
     *
     * @return 列宽
     */
    int width() default -1;

}
