package hht.dragon.office.exception;

/**
 * 不是Excel导入目标类.
 * User: huang
 * Date: 2018-8-18
 */
public class NotExcelModelException extends RuntimeException {

    private static final String MESSAGE = "Is not an export entity class of Excel";

    public NotExcelModelException() {
        super(MESSAGE);
    }

    public NotExcelModelException(Throwable cause) {
        super(MESSAGE, cause);
    }

    protected NotExcelModelException(Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(MESSAGE, cause, enableSuppression, writableStackTrace);
    }

}
