package hht.dragon.office.annotation.handler;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.annotation.ExcelDateType;
import hht.dragon.office.enums.DateFormatType;
import hht.dragon.office.exception.NotExcelModelException;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * 处理实体类注解的工具类.
 *
 * @author: huang
 * @Date: 18-10-31
 */
public class ExcelModelAnnotationUtil {
    private static ExcelModelAnnotationUtil ourInstance = new ExcelModelAnnotationUtil();

    public static ExcelModelAnnotationUtil getInstance() {
        return ourInstance;
    }

    private ExcelModelAnnotationUtil() {
    }

    /**
     * 是否通过指定列号读取数据.
     * @param cla 实体类
     * @return
     */
    public boolean isReadByColIndex(Class cla) {
        Excel excel = (Excel) cla.getAnnotation(Excel.class);
        if (excel == null) {
            throw new NotExcelModelException();
        }
        return excel.colIndex();
    }

    /**
     * 获取与excel文件对应的属性(通过注解指定的行)
     * @param cla 配置的实体类
     * @return
     */
    public Map<Integer, Field> readField(Class cla) {
        // 获取实体类的所有属性
        // TODO 得实验继承的类
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
     * 获取与excel文件对应的属性(通过标题行)
     * @param cla 配置的实体类
     * @param colTitle 读取excel文件后获取的与标题对应的map
     * @return
     */
    public Field[] readField(Class cla, Map<Integer, String> colTitle) {
        // 获取实体类的所有属性
        // TODO 要考虑继承
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
                if (anno.name().trim().equals(title.trim())) {
                    configField[i] = field;
                    break;
                }
            }
        }
        return configField;
    }

    /**
     * 获取属性的日期格式.
     * @param field 属性
     * @return 格式
     */
    public String getDateType(Field field) {
        ExcelDateType type = field.getAnnotation(ExcelDateType.class);
        if (type != null) {
            return type.value().getValue();
        }
        // TODO 其他格式
        return DateFormatType.DAY.getValue();
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
