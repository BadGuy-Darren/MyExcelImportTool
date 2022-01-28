package com.foxconn.indint.utils.getexcelutil.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel取值时，获取动态表头及本行（/列）对应表头栏位
 * 1. 注解作用于List类型字段
 * 2. titleRank表示表头所在行（/列），必填（横向表格代表行序，纵向表格表示列序）
 * 3. 从当前一直获取到最后一个表头及对应数据栏位
 * 4. 作用于实体类的最后一个字段，该字段后的其他字段将不会赋值
 * 5. enableDuplicateCheck复性检查：
 *      （1）依赖于实体类的equals方法的实现
 *      （2）会一定程度影响取值效率（看数据量），非必要不开启
 */
@Target(value = ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
@Inherited
public @interface DynamicRank {
    int titleRank();
    boolean enableDuplicateCheck() default false;
}
