package hht.dragon.office.utils;

import hht.dragon.office.annotation.ExcelDateType;
import hht.dragon.office.enums.DateFormatType;

import java.lang.reflect.Field;

/**
 * 自定义注解解析工具类.
 * User: huang
 * Date: 2018-8-23
 */
public class AnnotationUtil {

    private static final AnnotationUtil util = new AnnotationUtil();

    private AnnotationUtil() {}

    public static AnnotationUtil getInstance() {
        return util;
    }

    /**
     * 读取属性注解的日期格式.
     * @param field
     * @return
     */
    public String getTimeType(Field field) {
        ExcelDateType anno = field.getAnnotation(ExcelDateType.class);
        if (anno == null) {
            return DateFormatType.DAY.getValue();
        }
        DateFormatType type = anno.value();
        return type.getValue();
    }

}
