package hht.dragon.office.excel;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.annotation.handler.ExcelModelAnnotationUtil;
import hht.dragon.office.exception.ExportException;
import hht.dragon.office.utils.ExportExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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

    private final static String LOW_VERSION_SUFFIX = ".xls";
    private final static String HEIGHT_VERSION_SUFFIX = ".xlsx";
    private final List<String> titles = new ArrayList<>();
    private String fileName;
    private ExportExcelUtil util;
    private ExcelModelAnnotationUtil modelUtil;
    private boolean isHeightVersion;

    public ExportExcel(boolean isHeightVersion) {
        util = ExportExcelUtil.getInstance();
        this.modelUtil = ExcelModelAnnotationUtil.getInstance();
        this.isHeightVersion = isHeightVersion;
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
        OutputStream out;
        Workbook workbook;
        if (isHeightVersion) {
            workbook = new XSSFWorkbook();
        } else {
            workbook = new HSSFWorkbook();
        }
        Sheet sheet = workbook.createSheet();
        Object val = values.get(0);
        List<Field> fields = getValueFields(val.getClass());
        int rowIndex = 0;
        boolean isreadTitle = true;
        for (int i = 0; i < values.size(); i++) {
            Row row = sheet.createRow(rowIndex);
            if (i == 0) {
                if (isreadTitle && !modelUtil.isReadByColIndex(val.getClass())) {
                    writeTitle(row);
                    i--;
                    rowIndex++;
                    isreadTitle = false;
                    continue;
                }
            }
            writeValue(row, fields, values.get(i));
            rowIndex++;
        }

        String exportFileName;
        if (isHeightVersion) {
            exportFileName = fileName + HEIGHT_VERSION_SUFFIX;
        } else {
            exportFileName = fileName + LOW_VERSION_SUFFIX;
        }
        File file = new File(exportFileName);
        if (!file.exists()) {
            file.createNewFile();
        }
        out = new FileOutputStream(file);
        workbook.write(out);

        out.close();
        workbook.close();
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
            // TODO 继承时
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
    private void writeTitle(Row row) {
        for (int i = 0; i < titles.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(titles.get(i));
        }
    }

    /**
     * 将实体属性的值写入excel中
     * @param row
     * @param fields
     * @param value
     */
    private void writeValue(Row row, List<Field> fields, Object value) throws IllegalAccessException {
        for (int j = 0; j < fields.size(); j++) {
            Field field = fields.get(j);
            field.setAccessible(true);
            Object val = field.get(value);
            Cell cell = row.createCell(j);
            util.writeValue(val, field, cell);
        }
    }

}
