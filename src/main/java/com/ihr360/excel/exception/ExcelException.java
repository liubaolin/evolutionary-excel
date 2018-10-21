package com.ihr360.excel.exception;

/**
 * @author richey
 */
public class ExcelException extends RuntimeException {

    private static final long serialVersionUID = -3469867174279134372L;

    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }

    public ExcelException(String message, Throwable cause) {
        super(message,cause);
    }


}
