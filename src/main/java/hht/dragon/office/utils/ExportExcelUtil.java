package hht.dragon.office.utils;

import org.apache.poi.hssf.usermodel.HSSFCell;

import java.lang.reflect.Field;
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
     * @param field 属性
     * @param value 日期
     * @return
     */
    public String dateToStr(Field field, Object value) {
        if (value instanceof Date) {
            AnnotationUtil util = AnnotationUtil.getInstance();
            String format = util.getTimeType(field);
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return sdf.format(value);
        }
        return String.valueOf(value);
    }

    public void writeValue(Object obj, Field field, HSSFCell cell) {
        if (obj instanceof Date) {
            String time = dateToStr(field, obj);
            cell.setCellValue(time);
            return;
        }
        Class cla = obj.getClass();
        ConvertUtil convertUtil = ConvertUtil.getInstance();
        if (cla == Integer.class || cla == int.class) {
            Integer value = convertUtil.convertInt(obj);
            cell.setCellValue(value);
            return;
        }
        if (cla == Double.class) {
            Double value = convertUtil.convertDouble(obj);
            cell.setCellValue(value);
            return;
        }
        String value = String.valueOf(obj);
        cell.setCellValue(value);
    }

}
