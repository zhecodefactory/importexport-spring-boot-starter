package com.wz.annotation;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author 王哲
 * @ClassName ExcelStyle
 * @create 2023--五月--下午9:11
 * @Description
 * @Version V1.0
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelStyle {
    String fontName() default "Arial";   //设置字体的名称，默认为Arial。
    short fontSize() default 12;         //设置字体的大小，默认为12。
    short fontColor() default 8;         //设置字体颜色，默认为8（黑色）。
    short cellColor() default 9;         //设置单元格背景颜色，默认为9（灰色）。
    boolean wrapText() default false;    //设置是否自动换行，默认为不换行。
    byte alignment() default 1;          //设置单元格文字水平对齐方式，默认为居中。
    byte verticalAlignment() default 1;  //设置单元格文字垂直对齐方式，默认为居中。
}
