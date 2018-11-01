package hht.dragon.office.utils;

import hht.dragon.office.annotation.handler.ExcelModelAnnotationUtil;
import org.apache.poi.ss.usermodel.Cell;

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
            ExcelModelAnnotationUtil util = ExcelModelAnnotationUtil.getInstance();
            String format = util.getTimeType(field);
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return sdf.format(value);
        }
        return String.valueOf(value);
    }

    /**
     * 向单元格内写如数据.
     * @param obj 需写入的数据
     * @param field 写入数据的属性信息
     * @param cell 单元格
     */
    public void writeValue(Object obj, Field field, Cell cell) {
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
