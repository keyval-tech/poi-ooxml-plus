package com.kovizone.poi.ooxml.plus.anno;

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
@Retention(RetentionPolicy.RUNTIME)
public @interface Criteria {

    /**
     * 条件集合<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     * 否则没有条件
     *
     * @return 条件集
     */
    String[] value();

}
