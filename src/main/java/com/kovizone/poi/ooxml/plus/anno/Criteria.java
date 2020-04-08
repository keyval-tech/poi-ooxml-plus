package com.kovizone.poi.ooxml.plus.anno;

import com.kovizone.poi.ooxml.plus.processor.CriteriaProcessors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 输出到单元格的条件
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Processor(CriteriaProcessors.class)
@Retention(RetentionPolicy.RUNTIME)
public @interface Criteria {

    /**
     * 显示条件表达式<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     * 若为false则值为null
     *
     * @return 条件集
     */
    String value();

}
