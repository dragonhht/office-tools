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
@Excel(name = "student", colIndex = true)
@Data
public class ExlelModel {

    @ExcelColumn(index = 0)
    private String studentId;
    @ExcelColumn(index = 2)
    private String studentClass;
    @ExcelColumn(index = 1)
    private String studentName;
    @ExcelColumn(index = 5)
    @ExcelDateType(DateFormatType.DAY)
    private Date date;
    @ExcelColumn(index = 3)
    private String qq;
    @ExcelColumn(index = 4)
    private double score;
}
