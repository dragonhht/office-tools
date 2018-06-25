package hht.dragon.office.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 作用于实体类上，表示为一个excel，sheet对象.
 * User: huang
 * Date: 2018/6/22
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    /** 导出Excel文件是的文件名. */
    String name() default "";

}
