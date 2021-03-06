package hht.dragon.office.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel中的列.
 * User: huang
 * Date: 2018/6/22
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /** 列标题. */
    String name() default "";
    /** 属性所在的列号，第一列为0，第二列为1，依次类推. */
    int index() default -1;

}
