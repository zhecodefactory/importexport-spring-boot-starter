package com.wz.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author 王哲
 * @ClassName ExcelField
 * @create 2023--五月--下午9:12
 * @Description
 * @Version V1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * 列名称或标题
     *
     * @return 列名称或标题
     */
    String title();


    /**
     * 导出时的排列顺序
     *
     * @return 排列顺序
     */
    int order() default 0;


}
