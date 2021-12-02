package com.darren.tools.getexcelutil.annotation;

import java.lang.annotation.Target;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Documented;
import java.lang.annotation.Inherited;

/**
 * excel取值时，限制值范围
 */
@Target(value = ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface ValueLimit {
    String[] limit() default {};
    String message() default "com.darren.tools.getexcelutil.annotation.ValueLimit.message()";
}
