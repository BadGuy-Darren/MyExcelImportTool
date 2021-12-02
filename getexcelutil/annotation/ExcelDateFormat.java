package com.darren.tools.getexcelutil.annotation;

import java.lang.annotation.Target;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Documented;
import java.lang.annotation.Inherited;

/**
 * excel取值时，指定日期格式，无法转化指定格式时会抛异常
 */
@Target(value = ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface ExcelDateFormat {
    String pattern() default "yyyy/MM/dd";
    String message() default "annotation.getExcelUtil.DateFormat.message";
}
