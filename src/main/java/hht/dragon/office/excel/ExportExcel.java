package hht.dragon.office.excel;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.exception.ExportException;
import hht.dragon.office.exception.NotExcelModelException;
import hht.dragon.office.utils.ExportExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * 导出Excel.
 * User: huang
 * Date: 2018/6/26
 */
public class ExportExcel {

    private final List<String> titles = new ArrayList<>();
    private String fileName;
    private ExportExcelUtil util;

    public ExportExcel() {
        util = ExportExcelUtil.getInstance();
    }

    /**
     * 导出excel.
     * @param values 装有需导出数据的实体的列表.
     */
    public void exportValues(List values) throws ExportException, IOException, IllegalAccessException {
        if (values == null) {
            throw new ExportException("导出的实体类数据不能为空");
        }
        if (values.size() < 1) {
            throw new ExportException("导出的实体类数据不能为空");
        }
        OutputStream out = null;
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet();
        Object val = values.get(0);
        List<Field> fields = getValueFields(val.getClass());
        for (int i = 0; i < values.size(); i++) {
            HSSFRow row = sheet.createRow(i);
            if (i == 0) {
                if (!isReadByColIndex(val.getClass())) {
                    writeTitle(row);
                    continue;
                }
            }
            writeValue(row, fields, values.get(i));
        }

        File file = new File(fileName +".xls");
        if (!file.exists()) {
            file.createNewFile();
        }
        out = new FileOutputStream(file);
        workbook.write(out);

        if (out != null) {
            out.close();
        }
        if (workbook != null) {
            workbook.close();
        }
    }

    /**
     * 是否通过指定列号读取数据.
     * @param cla 实体类
     * @return
     */
    private boolean isReadByColIndex(Class cla) {
        Excel excel = (Excel) cla.getAnnotation(Excel.class);
        if (excel == null) {
            throw new NotExcelModelException();
        }
        return excel.colIndex();
    }

    /**
     * 获取文件名、标题行，并返回需要读取的属性.
     * @param cla excel实体类
     * @return 需要读取的属性
     */
    private List<Field> getValueFields(Class cla) {
        List<Field> values = new ArrayList<>();
        Excel excel = (Excel) cla.getAnnotation(Excel.class);
        if (excel != null) {
            fileName = excel.name();
            Field[] fields = cla.getDeclaredFields();
            for (Field field : fields) {
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                if (column != null) {
                    titles.add(column.name());
                    values.add(field);
                }
            }
        } else {
            throw new ExportException("实体类未注明为excel实体类");
        }
        return values;
    }

    /**
     * 写入标题行。
     * @param row
     */
    private void writeTitle(HSSFRow row) {
        for (int i = 0; i < titles.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(titles.get(i));
        }
    }

    /**
     * 将实体属性的值写入excel中
     * @param row
     * @param fields
     * @param value
     */
    private void writeValue(HSSFRow row, List<Field> fields, Object value) throws IllegalAccessException {
        for (int j = 0; j < fields.size(); j++) {
            Field field = fields.get(j);
            field.setAccessible(true);
            Object val = field.get(value);
            HSSFCell cell = row.createCell(j);
            util.writeValue(val, field, cell);
        }
    }

}
