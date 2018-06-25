package hht.dragon.office.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * 作用于实体类上，表示为一个excel，sheet对象.
 * User: huang
 * Date: 2018/6/22
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    /** 导出Excel文件是的文件名. */
    String name() default "";

}
