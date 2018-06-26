package hht.dragon.office.utils;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 导出excel文件的工具类.
 * User: huang
 * Date: 2018/6/26
 */
public class ExportExcelUtil {

    private static final ExportExcelUtil util = new ExportExcelUtil();

    private ExportExcelUtil() {

    }

    public static ExportExcelUtil getInstance() {
        return util;
    }

    /**
     * 将日期格式转换为一定的日期格式.
     * @param format 日期格式
     * @param value 日期
     * @return
     */
    public String dateToStr(String format, Object value) {
        if (value instanceof Date) {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return sdf.format(value);
        }
        return String.valueOf(value);
    }

}
