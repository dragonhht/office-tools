package hht.dragon.office.base;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.excel.ExportExcel;
import hht.dragon.office.excel.ImportExcel;
import hht.dragon.office.utils.ExportExcelUtil;
import hht.dragon.office.utils.ReadExcelConfigUtil;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Description.
 * User: huang
 * Date: 2018/6/22
 */
public class TestExcel {

    @Test
    public void testimport() throws IOException, IllegalAccessException {
        long start = System.currentTimeMillis();
        File file = new File("test.xls");
        InputStream input = new FileInputStream(file);

        ImportExcel excel = new ImportExcel();
        List<ExlelModel> values = new ArrayList();
        excel.importValue(input, 0, ExlelModel.class, values);

        input.close();

        long end = System.currentTimeMillis();
        System.out.println("运行时间：" + (end - start));

        ExportExcel export = new ExportExcel();
        export.exportValues(values, ".");
    }

    @Test
    public void test1(){
        Object v = 6;
//        ReadExcelConfigUtil util = ReadExcelConfigUtil.getInstance();
//        System.out.println(util.convertValue(v, String.class));
        String str = "6.9";
        System.out.println(str.substring(0,str.indexOf('.')));

    }

    @Test
    public void test2() {
        Object o = new ExlelModel();
        ExportExcelUtil util = ExportExcelUtil.getInstance();
    }

}
