package hht.dragon.office.enums;

/**
 * 日期格式.
 * User: huang
 * Date: 2018-8-17
 * @author huang
 */
public enum DateFormatType {
    /** yyyy-MM-dd .*/
    DAY("yyyy-MM-dd"),
    /** yyyy-MM-dd HH:mm:ss .*/
    TIMES("yyyy-MM-dd HH:mm:ss");

    private String value;

    DateFormatType(String format) {
        value = format;
    }

    public String getValue() {
        return value;
    }

}
