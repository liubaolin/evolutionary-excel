package com.ihr360.excel.metaData;

import com.fasterxml.jackson.annotation.JsonInclude;

import java.io.Serializable;

/**
 * @author Stone.Shi
 * @description
 * @date 2018-04-13 09:22:39.
 */
@JsonInclude(JsonInclude.Include.NON_NULL)
public class ExcelHeaderCellEntity implements Serializable {
    private static final long serialVersionUID = -884789181226424879L;

    public static final String TYPE_TEXT = "text";
    public static final String TYPE_NUMBERIC = "numeric";
    public static final String TYPE_DATE = "date";
    public static final String TYPE_TIME = "time";

    /**
     * 字段属性名称（表头可显示的名称）
     */
    private String fieldName;

    /**
     * 字段名称
     */
    private String columnName;

    /**
     * 字段类型
     */
    private String type;

    private boolean readOnly = true;
    /**
     * 是否必填
     */
    private Boolean isRequired;
    /**
     * 是否唯一
     */
    private Boolean isUnique;
    /**
     * 是否匹配
     */
    private Boolean isMatching;
    /**
     * 字段最小长度
     */
    private String minLength;
    /**
     * 字段最大长度
     */
    private String maxLength;
    /**
     * 字段长度
     */
    private Integer length;
    /**
     * 正则表达式
     */
    private String regexp;

    public ExcelHeaderCellEntity() {
    }

    public ExcelHeaderCellEntity(String fieldName, String columnName) {
        this.fieldName = fieldName;
        this.columnName = columnName;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public boolean getReadOnly() {
        return readOnly;
    }

    public void setReadOnly(boolean readOnly) {
        this.readOnly = readOnly;
    }

    public Boolean getRequired() {
        return isRequired;
    }

    public void setRequired(Boolean required) {
        isRequired = required;
    }

    public Boolean getUnique() {
        return isUnique;
    }

    public void setUnique(Boolean unique) {
        isUnique = unique;
    }

    public Boolean getMatching() {
        return isMatching;
    }

    public void setMatching(Boolean matching) {
        isMatching = matching;
    }

    public String getMinLength() {
        return minLength;
    }

    public void setMinLength(String minLength) {
        this.minLength = minLength;
    }

    public String getMaxLength() {
        return maxLength;
    }

    public void setMaxLength(String maxLength) {
        this.maxLength = maxLength;
    }

    public Integer getLength() {
        return length;
    }

    public void setLength(Integer length) {
        this.length = length;
    }

    public String getRegexp() {
        return regexp;
    }

    public void setRegexp(String regexp) {
        this.regexp = regexp;
    }
}
