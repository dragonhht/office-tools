package hht.dragon.office.base;

import hht.dragon.office.annotation.Excel;
import hht.dragon.office.annotation.ExcelColumn;
import hht.dragon.office.annotation.ExcelDateType;
import hht.dragon.office.enums.DateFormatType;
import lombok.Data;

import java.util.Date;

/**
 * 学生.
 * User: huang
 * Date: 2018/6/22
 */
@Excel(name = "student")
@Data
public class ExlelModel {

    @ExcelColumn(name = "学号")
    private String studentId;
    @ExcelColumn(name = "班级")
    private String studentClass;
    @ExcelColumn(name = "姓名")
    private String studentName;
    @ExcelColumn(name = "日期")
    @ExcelDateType(DateFormatType.DAY)
    private Date date;
    @ExcelColumn(name = "qq")
    private String qq;
    @ExcelColumn(name = "成绩")
    private double score;
}
