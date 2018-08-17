package hht.dragon.office.annotation;

import hht.dragon.office.enums.DateFormatType;

/**
 * Excel的日期格式.
 * User: huang
 * Date: 2018-8-17
 */
public @interface ExcelDateType {

    /** 日期格式。 */
    DateFormatType value() default DateFormatType.DAY;

}
