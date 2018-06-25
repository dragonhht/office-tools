package hht.dragon.office.base;

import hht.dragon.office.excel.ImportExcel;
import hht.dragon.office.utils.ReadExcelConfigUtil;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
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
    public void testimport() throws IOException {
        File file = new File("test.xls");
        InputStream input = new FileInputStream(file);

        ImportExcel excel = new ImportExcel();
        List<ExlelModel> values = new ArrayList();
        excel.importValue(input, 0, ExlelModel.class, values);

        input.close();

        for (ExlelModel obj : values) {
            System.out.println(obj);
            System.out.println(obj.getDate());
        }
    }

    @Test
    public void test1(){
        Object v = 6;
//        ReadExcelConfigUtil util = ReadExcelConfigUtil.getInstance();
//        System.out.println(util.convertValue(v, String.class));
        String str = "6.9";
        System.out.println(str.substring(0,str.indexOf('.')));

    }
}
