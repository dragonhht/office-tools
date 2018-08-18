package hht.dragon.office.utils;

import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.annotation.ExcelDateType;
import hht.dragon.office.enums.DateFormatType;

import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 读取excel实体配置的工具类.
 * Date: 2018/6/25
 * @author huang
 */
public class ReadExcelConfigUtil {
    private static final ReadExcelConfigUtil util = new ReadExcelConfigUtil();

    private static final String EXCEL_BIG_NUM_FLAG = "E";

    private ReadExcelConfigUtil() {}

    public static ReadExcelConfigUtil getInstance() {
        return util;
    }

    /**
     * 获取与excel文件对应的属性(通过标题行)
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
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 获取与excel文件对应的属性(通过注解制定的行)
     * @param cla 配置的实体类
     * @return
     */
    public Map<Integer, Field> readField(Class cla) {
        // 获取实体类的所有属性
        Field[] fields = cla.getDeclaredFields();
        Map<Integer, Field> map = new HashMap<>(5);
        for (Field field : fields) {
            ExcelColumn anno = field.getAnnotation(ExcelColumn.class);
            if (anno == null) {
                continue;
            }
            if (anno.index() != -1) {
                map.put(anno.index(), field);
            }
        }
        return map;
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
                Object o = convertValue(values[i], fields[i]);
                fields[i].set(obj, o);
            }
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        return obj;
    }

    /**
     * 将数据转换为指定的类型.
     * @param o 数据
     * @param field 属性
     * @return 转换后的数据
     */
    private Object convertValue(Object o, Field field) {
        Class cla = field.getType();
        if (cla == String.class) {
            String dateFormat = getDateType(field);
            return convertString(o, dateFormat);
        }
        if (cla == Date.class) {
            String dateFormat = getDateType(field);
            return convertDate(o, dateFormat);
        }
        if ( (cla == Integer.class || cla == int.class)) {
            return convertInt(o);
        }
        if ((cla == Double.class || cla == double.class)) {
            return convertDouble(o);
        }
        if ((cla == Float.class || cla == float.class)) {
            return convertFloat(o);
        }
        if (!(o instanceof Number) && !(o instanceof String)) {
            return cla.cast(o);
        } else {
            return null;
        }
    }

    /**
     * 转换为字符串.
     * @param o 数据
     * @param dateFormat 日期类型
     * @return 转换后的数据
     */
    private String convertString(Object o, String dateFormat) {

        // 将数值转换为字符串
        if (o instanceof Number || o instanceof Boolean) {
            String str = String.valueOf(o);
            if (str.contains(EXCEL_BIG_NUM_FLAG)) {
                return String.valueOf(new DecimalFormat("#").format(o));
            }
            return String.valueOf(o);
        }
        // 将日期转换为字符串
        if (o instanceof Date) {
            SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
            return sdf.format(o);
        }
        if (o instanceof String) {
            return String.valueOf(o);
        }
        return null;
    }

    /**
     * 转换为日期.
     * @param o 数据
     * @param dataFormat 日期格式
     * @return 转换后的数据
     */
    private Date convertDate(Object o , String dataFormat) {
        if (o instanceof String) {
            SimpleDateFormat sdf = new SimpleDateFormat(dataFormat);
            try {
                return sdf.parse(String.valueOf(o));
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        if (o instanceof Date) {
            return (Date) o;
        }
        return null;
    }

    /**
     * 转换为整型.
     * @param o 数据
     * @return 转换后的数据
     */
    private Integer convertInt(Object o) {
        if (o == null) {
            return 0;
        }
        String str = String.valueOf(o);
        int index = str.indexOf('.');
        if (index > -1) {
            str = str.substring(0, index);
        }
        return Integer.parseInt(str);
    }

    /**
     * 转换为双精度浮点型.
     * @param o 数据
     * @return 转换后的数据
     */
    private Double convertDouble(Object o) {
        if (o == null) {
            return 0.0;
        }
        return Double.parseDouble(String.valueOf(o));
    }

    /**
     * 转换为单精度浮点型.
     * @param o 数据
     * @return 转换后的数据
     */
    private Float convertFloat(Object o) {
        if (o == null) {
            return 0.0f;
        }
        return Float.parseFloat(String.valueOf(o));
    }

    /**
     * 获取属性的日期格式.
     * @param field 属性
     * @return 格式
     */
    private String getDateType(Field field) {
        ExcelDateType type = field.getAnnotation(ExcelDateType.class);
        if (type != null) {
            return type.value().getValue();
        }
        return DateFormatType.DAY.getValue();
    }

}
