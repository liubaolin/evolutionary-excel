package com.ihr360.excel.core.metaData;

import java.io.Serializable;

/**
 * describe:
 * Excel行数据信息
 *
 * @author Monkey
 * @date 18-3-15
 */
public class ExcelDataCellEntity implements Serializable {

    private static final long serialVersionUID = 3359665481256125799L;

    private String columnName;

    /**
     * execl中的值
     */
    private String value;

    /**
     * 数据库存在的值
     */
    private String originalValue;

    /**
     * 错误类型
     */
    private String errorType;

    public ExcelDataCellEntity() {
    }

    public ExcelDataCellEntity(String columnName, String value) {
        this(columnName, value, value);
    }

    public ExcelDataCellEntity(String columnName, String value, String originalValue) {
        this.columnName = columnName;
        this.value = value;
        this.originalValue = originalValue;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getOriginalValue() {
        return originalValue;
    }

    public void setOriginalValue(String originalValue) {
        this.originalValue = originalValue;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getErrorType() {
        return errorType;
    }

    public void setErrorType(String errorType) {
        this.errorType = errorType;
    }
}
