package hht.dragon.office.utils;

import hht.dragon.office.annotation.ExcelColumn;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

/**
 * 读取excel实体配置的工具类.
 * User: huang
 * Date: 2018/6/25
 */
public class ReadExcelConfigUtil {
    private static final ReadExcelConfigUtil util = new ReadExcelConfigUtil();

    private ReadExcelConfigUtil() {}

    public static ReadExcelConfigUtil getInstance() {
        return util;
    }

    /**
     * 获取与excel文件对应的属性
     * @param cla 配置的实体类
     * @param colTitle 读取excel文件后获取的与标题对应的map
     * @return
     */
    public Field[] readField(Class cla, Map<Integer, String> colTitle) {
        try {
            Object obj = cla.newInstance();
            // 获取实体类的所有属性
            Field[] fields = cla.getDeclaredFields();
            int len = colTitle.size();
            // 按标题行顺序依次存放实体属性
            Field[] configField = new Field[len];
            for (int i = 0; i < len; i++) {
                String title = colTitle.get(i);
                for (Field field : fields) {
                    ExcelColumn anno = field.getAnnotation(ExcelColumn.class);
                    if (anno == null) {
                        continue;
                    }
                    // 判断注解name属性的值与map中的是否对应
                    if (anno.name().equals(title)) {
                        configField[i] = field;
                        break;
                    }
                }
            }
            return configField;
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 创建一个填充有从excel读取到数据的实体对象.
     * @param fields 需填充的属性域
     * @param modelCalss 实体类
     * @return
     */
    public Object newModel(Field[] fields, Class modelCalss, Object[] values) {
        Object obj = null;
        try {
            obj = modelCalss.newInstance();
            for (int i = 0; i < fields.length; i++) {
                fields[i].setAccessible(true);
                Object o = convertValue(values[i], fields[i].getType());
                fields[i].set(obj, o);
            }
        } catch (InstantiationException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return obj;
    }

    /**
     * 将数据转换为指定的类型.
     * @param o 数据
     * @param cla 类型
     * @return 转换后的数据
     */
    public Object convertValue(Object o, Class cla) {
        if (cla == String.class) {
            // 将数值转换为字符串
            if (o instanceof Number || o instanceof Boolean) {
                return String.valueOf(o);
            }
            // 将日期转换为字符串
            if (o instanceof Date) {
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                return sdf.format(o);
            }
        }
        if ( (cla == Integer.class || cla == int.class)) {
            if (o == null) {
                return 0;
            }
            return Integer.parseInt(String.valueOf(o));
        }
        if ((cla == Double.class || cla == double.class)) {
            if (o == null) {
                return 0.0;
            }
            return Double.parseDouble(String.valueOf(o));
        }
        if ((cla == Float.class || cla == float.class)) {
            if (o == null) {
                return 0.0;
            }
            return Float.parseFloat(String.valueOf(o));
        }
        if (!(o instanceof Number) && !(o instanceof String)) {
            return cla.cast(o);
        } else {
            return null;
        }
    }

}
