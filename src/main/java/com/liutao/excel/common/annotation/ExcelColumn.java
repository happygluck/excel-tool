package com.liutao.excel.common.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Excel列标识
 */
@Target({METHOD, FIELD})
@Retention(RUNTIME)
public @interface ExcelColumn {
    /**
     * 列顺序
     *
     * @return
     */
    int index() default 0;

    /**
     * 列标题
     *
     * @return
     */
    String name() default "";

    /**
     * 列宽（一个汉字占2个宽度）
     *
     * @return
     */
    int width() default 10;
}
