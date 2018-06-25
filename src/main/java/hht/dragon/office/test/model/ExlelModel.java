package hht.dragon.office.test.model;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import lombok.Data;

import java.util.Date;

/**
 * Description.
 * User: huang
 * Date: 2018/6/22
 */
@Excel(name = "test1")
@Data
public class ExlelModel {

    @ExcelColumn(name = "名称")
    private double title;
    @ExcelColumn(name = "测试")
    private double test;
    @ExcelColumn(name = "qq")
    private double qq;
    @ExcelColumn(name = "日期")
    private Date date;

    private String name;

    @Override
    public String toString() {
        return "ExlelModel{" +
                "title=" + title +
                ", test=" + test +
                ", qq=" + qq +
                ", date=" + date +
                ", name='" + name + '\'' +
                '}';
    }
}