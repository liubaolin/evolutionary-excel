package com.ihr360.excel.commons.exception;

public class ExcelDateParseException extends RuntimeException {
    private static final long serialVersionUID = -8702493823845839194L;

    public ExcelDateParseException() {
        super();
    }

    public ExcelDateParseException(String message) {
        super(message);
    }

    public ExcelDateParseException(Throwable cause) {
        super(cause);
    }

    public ExcelDateParseException(String message, Throwable cause) {
        super(message, cause);
    }

}
