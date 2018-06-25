package hht.dragon.office.excel;

import hht.dragon.office.utils.ReadExcelConfigUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellType;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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
     * @param row
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
            Field[] fields = null;

            int len = sheet.getLastRowNum();
            for (int i = 0; i < len; i++) {
                HSSFRow row = sheet.getRow(i);
                if (i == 0) {
                    setColTitle(row);
                    fields = util.readField(modelCalss, colTitle);
                    continue;
                }
                Object[] vals = getValues(row);
                values.add(util.newModel(fields, modelCalss, vals));
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

}
