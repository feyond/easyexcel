package com.github.feyond.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Sheet 注解定义
 *
 * @author
 * @create 2017-09-26 11:30
 **/
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {

    /**
     * 导出Sheet名
     */
    String value();
}
