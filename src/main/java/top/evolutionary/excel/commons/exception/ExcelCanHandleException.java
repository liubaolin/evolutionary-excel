package top.evolutionary.excel.commons.exception;

/**
 * @author richey
 */
public class ExcelCanHandleException extends RuntimeException {
    private static final long serialVersionUID = -3755477657584978882L;

    private Object[] args;

    public ExcelCanHandleException(ExcellExceptionType excellExceptionType) {
        super(excellExceptionType.key,new ExcelException(excellExceptionType.name));
    }


    public ExcelCanHandleException(ExcellExceptionType excellExceptionType,Object[] args) {
        super(excellExceptionType.key,new ExcelException(excellExceptionType.name));
        this.args = args;
    }

    public Object[] getArgs() {
        return args;
    }

}
