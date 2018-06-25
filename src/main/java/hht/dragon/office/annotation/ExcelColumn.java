package hht.dragon.office.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * excel中的列.
 * User: huang
 * Date: 2018/6/22
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {

    /** 列标题. */
    String name() default "";

}
