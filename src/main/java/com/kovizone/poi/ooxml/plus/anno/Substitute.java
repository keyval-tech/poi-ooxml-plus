package com.kovizone.poi.ooxml.plus.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定在特殊条件下的替补字段
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Substitute {

    /**
     * 替补条件
     *
     * @return 替补条件
     */
    String[] condition() default {"IS NUll"};

    /**
     * 替补值<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     * 否则为常量
     *
     * @return 替补属性名
     */
    String value();
}
