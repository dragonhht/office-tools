package hht.dragon.office.utils;

import java.lang.reflect.Field;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 类型转换工具类.
 * User: huang
 * Date: 18-8-26
 */
public class ConvertUtil {

    private static final ConvertUtil util = new ConvertUtil();

    private static final String EXCEL_BIG_NUM_FLAG = "E";

    private ConvertUtil() {

    }

    public static ConvertUtil getInstance() {
        return util;
    }

    /**
     * 转换为字符串.
     * @param o 数据
     * @param dateFormat 日期类型
     * @return 转换后的数据
     */
    public String convertString(Object o, String dateFormat) {

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
    public Date convertDate(Object o , String dataFormat) {
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
    public Integer convertInt(Object o) {
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
    public Double convertDouble(Object o) {
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
    public Float convertFloat(Object o) {
        if (o == null) {
            return 0.0f;
        }
        return Float.parseFloat(String.valueOf(o));
    }

}
