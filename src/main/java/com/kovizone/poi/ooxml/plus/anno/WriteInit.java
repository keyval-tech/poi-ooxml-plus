package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.ExcelWriter;
import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteInitProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author <a href="mailto:kovichen@163.com">KoviChen</a>
 * @version 1.0
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Processor(WriteInitProcessors.class)
public @interface WriteInit {

    /**
     * Sheet标签名
     *
     * @return Sheet标签名
     * @see ExcelWriter#SHEET_NUM 页码占位符
     */
    String sheetName() default "Sheet" + ExcelWriter.SHEET_NUM;

    /**
     * 默认行高
     *
     * @return 默认行高
     */
    short rowHeight() default -1;

    /**
     * 默认列宽
     *
     * @return 默认列宽
     */
    int columnWidth() default -1;

}
