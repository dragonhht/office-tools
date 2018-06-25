package hht.dragon.office.base;

import hht.dragon.office.excel.ImportExcel;
import hht.dragon.office.test.model.ExlelModel;
import hht.dragon.office.utils.ReadExcelConfigUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Description.
 * User: huang
 * Date: 2018/6/22
 */
public class TestExcel {

    @Test
    public void testimport() throws IOException {
        File file = new File("test.xls");
        InputStream input = new FileInputStream(file);

        ImportExcel excel = new ImportExcel();
        List values = new ArrayList();
        excel.importValue(input, 0, ExlelModel.class, values);

        input.close();

        for (Object obj : values) {
            System.out.println(obj);
        }
    }
}
