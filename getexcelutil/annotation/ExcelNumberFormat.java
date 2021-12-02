package com.darren.tools.getexcelutil.annotation;

import java.lang.annotation.Target;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Documented;
import java.lang.annotation.Inherited;

/**
 * excel取值时，指定数字格式，无法转化指定格式时会抛异常
 *  when例子：when = "field == value"
 */
@Target(value = ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface ExcelNumberFormat {
    String format() default "#.##";
    // "field == value"
    String when() default "";
    String message() default "annotation.getExcelUtil.annotation.ExcelNumberFormat.message";
}
