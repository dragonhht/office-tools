package hht.dragon.office.base;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.annotation.ExcelDateType;
import lombok.Data;

import java.util.Date;

/**
 * Description.
 * User: huang
 * Date: 2018/6/22
 */
@Excel(name = "test1", colIndex = true)
@Data
public class ExlelModel {

    @ExcelColumn(name = "名称")
    private int title;
    @ExcelColumn(name = "测试")
    private double test;
    @ExcelColumn(name = "qq")
    private String qq;
    @ExcelColumn(name = "日期")
    @ExcelDateType
    private Date date;

    private String name;

}
