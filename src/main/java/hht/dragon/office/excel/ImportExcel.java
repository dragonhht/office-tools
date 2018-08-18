package hht.dragon.office.excel;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.exception.NotExcelModelException;
import hht.dragon.office.utils.ReadExcelConfigUtil;
import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 导入Excel.
 * User: huang
 * Date: 2018/6/22
 */
public class ImportExcel {

    /** 用于保存列数与行标题的关系. */
    private final Map<Integer, String> colTitle = new HashMap<>();
    private ReadExcelConfigUtil util;

    public ImportExcel() {
        util = ReadExcelConfigUtil.getInstance();
    }


    /**
     * 保存列数与标题的关系.
     */
    private void setColTitle(HSSFRow row) throws NullPointerException {
        if (row != null) {
            int num = row.getLastCellNum();
            for (int i = 0; i < num; i++) {
                HSSFCell cell = row.getCell(i);
                // TODO 暂支持字符串
                colTitle.put(i, cell.getStringCellValue());
            }
        } else {
            throw new NullPointerException("HSSFRow不能为空");
        }
    }

    /**
     * 读取单元格的值.
     * @param cell
     * @return
     */
    private Object getValue(HSSFCell cell) {
        if (cell == null) {
            return null;
        }
        // TODO 考虑下还有那些类型的值
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
                return null;
            default:
                return null;
        }
    }

    /**
     * 读取一行excel数据.
     * @param row Excel中的行
     * @return
     */
    private Object[] getValues(HSSFRow row) {
        Object[] objects = null;
        if (row != null) {
            int len = row.getLastCellNum();
            objects = new Object[len];
            for (int i = 0; i < len; i++) {
                HSSFCell cell = row.getCell(i);
                objects[i] = getValue(cell);
            }
        }
        return objects;
    }

    /**
     * 读取一行Excel的数据(指定的列)
     * @param row Excel中的行
     * @param col 需读取的列
     * @return
     */
    private Object[] getValues(HSSFRow row, Set<Integer> col) {
        Object[] objects = null;
        if (row != null) {
            objects = new Object[col.size()];
            int colIndex = 0;
            System.out.println(col);
            for (int index : col) {
                System.out.println("col: " + index);
                HSSFCell cell = row.getCell(index);
                objects[colIndex] = getValue(cell);
                colIndex++;
            }
        }
        return objects;
    }

    /**
     * 读取excel文件的内容.
     * @param input 文件的输入流
     * @param sheetIndex sheet页码
     * @param modelCalss 接受每行内容的实体类
     * @param values 接收数据的列表
     */
    public void importValue(InputStream input, int sheetIndex, Class modelCalss, List values) {
        HSSFWorkbook workbook = null;
        try {
            workbook = new HSSFWorkbook(input);
            HSSFSheet sheet = workbook.getSheetAt(sheetIndex);
            if (isReadByColIndex(modelCalss)) {
                importValueByCol(sheet, modelCalss, values);
            } else {
                // 通过标题行
                importValueByTitle(sheet, modelCalss, colTitle, values);
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
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
     * 导入Excel的数据，通过标题行.
     * @param sheet excel中的sheet.
     * @param modelClass 装载数据的实体
     * @param colTitle 标题行
     */
    private void importValueByTitle(HSSFSheet sheet, Class modelClass, Map<Integer, String> colTitle,
                                    List values) {
        Field[] fields = null;
        int len = sheet.getLastRowNum();
        for (int i = 0; i < len; i++) {
            HSSFRow row = sheet.getRow(i);
            if (i == 0) {
                // TODO 若第一行不是标题行怎么办
                setColTitle(row);
                fields = util.readField(modelClass, colTitle);
                continue;
            }
            Object[] vals = getValues(row);
            values.add(util.newModel(fields, modelClass, vals));
        }
    }

    private void importValueByCol(HSSFSheet sheet, Class modelClass, List values) {
        Map<Integer, Field> map = util.readField(modelClass);
        Set<Integer> cols = map.keySet();
        Field[] fields = new Field[cols.size()];
        fields = map.values().toArray(fields);
        int len = sheet.getLastRowNum();
        for (int i = 0; i < len; i++) {
            HSSFRow row = sheet.getRow(i);
            Object[] vals = getValues(row, cols);
            values.add(util.newModel(fields, modelClass, vals));
        }
    }

    /**
     * 读取excel文件的内容.
     * @param file excel文件
     * @param sheetIndex sheet页码
     * @param modelCalss 接受每行内容的实体类
     * @param values 接收数据的列表
     */
    public void importValue(File file, int sheetIndex, Class modelCalss, List values) {
        InputStream input = null;
        try {
            input = new FileInputStream(file);
            importValue(input, sheetIndex, modelCalss, values);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } finally {
            if (input != null) {
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
