package hht.dragon.office.utils;

import hht.dragon.office.annotation.handler.ExcelModelAnnotationUtil;

import java.lang.reflect.Field;
import java.text.ParseException;
import java.util.Date;

/**
 * 读取excel实体配置的工具类.
 * Date: 2018/6/25
 * @author huang
 */
public class ReadExcelConfigUtil {
    private static final ReadExcelConfigUtil util = new ReadExcelConfigUtil();
    private static final ExcelModelAnnotationUtil modelUtil = ExcelModelAnnotationUtil.getInstance();

    private ReadExcelConfigUtil() {}

    public static ReadExcelConfigUtil getInstance() {
        return util;
    }

    /**
     * 创建一个填充有从excel读取到数据的实体对象.
     * @param fields 需填充的属性域
     * @param modelCalss 实体类
     * @return
     */
    public <T> T newModel(Field[] fields, Class<T> modelCalss, Object[] values) throws IllegalAccessException, InstantiationException, ParseException {
        T obj;
        obj = modelCalss.newInstance();
        for (int i = 0; i < fields.length; i++) {
            fields[i].setAccessible(true);
            Object o = convertValue(values[i], fields[i]);
            fields[i].set(obj, o);
        }
        return obj;
    }

    /**
     * 将数据转换为指定的类型.
     * @param o 数据
     * @param field 属性
     * @return 转换后的数据
     */
    private Object convertValue(Object o, Field field) throws ParseException {
        // TODO 还有其他其他类型吧
        ConvertUtil convertUtil = ConvertUtil.getInstance();
        Class cla = field.getType();
        if (cla == String.class) {
            String dateFormat = modelUtil.getDateType(field);
            return convertUtil.convertString(o, dateFormat);
        }
        if (cla == Date.class) {
            String dateFormat = modelUtil.getDateType(field);
            return convertUtil.convertDate(o, dateFormat);
        }
        if ( (cla == Integer.class || cla == int.class)) {
            return convertUtil.convertInt(o);
        }
        if ((cla == Double.class || cla == double.class)) {
            return convertUtil.convertDouble(o);
        }
        if ((cla == Float.class || cla == float.class)) {
            return convertUtil.convertFloat(o);
        }
        if (!(o instanceof Number) && !(o instanceof String)) {
            return cla.cast(o);
        } else {
            return null;
        }
    }

}
