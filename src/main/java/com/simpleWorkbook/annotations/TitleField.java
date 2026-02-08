package com.simpleWorkbook.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表格标题标注，代表sheet中的标题行属性
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface TitleField {

    /**
     * 标题名称
     */
    String value();

    /**
     * 字典值，会生成数据校验下拉
     */
    String[] dictValues() default {};

    String dictSheetName() default "";

    /**
     * 列宽
     */
    int colWidth() default 15;


    boolean listValuesInSingleCell() default false;

    String listValuesInSingleCellSplitter() default ",";
}
