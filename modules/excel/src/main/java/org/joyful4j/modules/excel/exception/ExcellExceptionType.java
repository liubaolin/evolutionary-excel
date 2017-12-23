package org.joyful4j.modules.excel.exception;

public enum  ExcellExceptionType {

    REPEATED_HEADER("excel_exception.enum.repeated_header", "存在重复的表头");

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
