package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.api.anno.Processor;
import com.kovizone.poi.ooxml.plus.processor.WriteInitProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定为Excel工作表对象
 *
 * @author KoviChen
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Processor(WriteInitProcessors.class)
public @interface WriteInit {

    /**
     * 指定显示工作簿名，
     * 使用#{sheetNum}占位将会替换为自动排序：1、2、3...，
     * 若没有#{sheetNum}，则强制在后面加入排序，
     *
     * @return 工作簿名
     */
    String sheetName() default "工作簿#{sheetNum}";

    /**
     * 默认单元格宽度，
     * 创建行后默认宽度会清除，
     *
     * @return 默认行宽
     */
    int defaultColumnWidth() default -1;

    /**
     * 默认行高度，
     * 创建行后默认高度会清除，
     *
     * @return 默认高度
     */
    short defaultRowHeight() default -1;
}