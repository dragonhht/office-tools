package hht.dragon.office.utils;

/**
 * 自定义注解解析工具类.
 * User: huang
 * Date: 2018-8-23
 */
public class AnnotationUtil {

    private static final AnnotationUtil util = new AnnotationUtil();

    private AnnotationUtil() {}

    private static AnnotationUtil getInstance() {
        return util;
    }

}
