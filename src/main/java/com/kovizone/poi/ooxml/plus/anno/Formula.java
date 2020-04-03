package com.kovizone.poi.ooxml.plus.anno;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 函数<BR/>
 *
 * @author KoviChen
 * @deprecated 暂未开发处理器
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Deprecated
public @interface Formula {

    /**
     * 保存函数<BR/>
     * 可用以下参数占位符<BR/>
     *
     * @return 函数
     */
    String formula();
}
