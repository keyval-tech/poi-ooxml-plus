package com.kovizone.poi.ooxml.plus.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 指定后缀属性
 *
 * @author KoviChen
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface Suffix {

    /**
     * 输出值拼接后缀<BR/>
     * 使用#[FieldName]指定属性的值<BR/>
     *
     * @return 拼接后缀集
     */
    String[] value();

}
