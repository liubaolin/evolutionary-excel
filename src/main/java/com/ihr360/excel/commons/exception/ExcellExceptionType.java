package com.ihr360.excel.commons.exception;

public enum  ExcellExceptionType {

    REPEATED_HEADER("excel_exception.enum.repeated_header", "存在重复的表头"),
    HEADER_DATA_TYPE_NOTSUPPORT("excel_exception.enum.header_data_type_not_support", "不支持当前数据类型的表头"),
    HEADER_DATE_FIELD_NOTSUPPORT("excel_exception.enum.header_data_type_not_support", "不支持日期类型的表头");

    String key;
    String name;

    private ExcellExceptionType(String key, String name) {
        this.key = key;
        this.name = name;
    }

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
