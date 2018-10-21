package com.ihr360.excel.specification;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author richey
 * 用于说明列的规则，如数据类型，是否必填
 *
 * 目前只用于javaBean的flexField弹性字段以及map结构的导入数据
 * 暂不支持javaBean自身的field，javaBean自身的field暂时只用注解配置
 */
public class ColumnSpecification implements Serializable{

    /**
     * 默认为false
     * 如果为true,则该规则用于除当前Columns之外的Column
     */
    private boolean ignoreColumn = false;

    /**
     * 单元格格式
     */
    private Class cellType;

    /**
     * 是否允许为空
     */
    private boolean allowNull = true;

    /**
     * 适用于那些列
     * 表头集合
     */
    private List<String> columns = new ArrayList<>();

    public ColumnSpecification() {
    }

    public ColumnSpecification(Class cellType, boolean allowNull) {
        this.cellType = cellType;
        this.allowNull = allowNull;
    }

    public Class getCellType() {
        return cellType;
    }

    public void setCellType(Class cellType) {
        this.cellType = cellType;
    }

    public List<String> getColumns() {
        return columns;
    }

    public boolean isAllowNull() {
        return allowNull;
    }

    public void setAllowNull(boolean allowNull) {
        this.allowNull = allowNull;
    }

    public boolean getIgnoreColumn() {
        return ignoreColumn;
    }

    public void setIgnoreColumn(boolean ignoreColumn) {
        this.ignoreColumn = ignoreColumn;
    }

    @Deprecated
    public void addColumns(String... columns) {
        if (columns == null || columns.length < 1) {
            return;
        }
        List<String> columnsList =  Arrays.stream(columns)
                .filter(column -> StringUtils.isNotBlank(column))
                .collect(Collectors.toList());
        if (CollectionUtils.isNotEmpty(columnsList)) {
            this.columns.addAll(columnsList);
        }

    }

    public void setColumns(List<String> columns) {
        this.columns = columns;
    }

    @Deprecated
    public static ColumnSpecification createCellSpecification(Class cellType, boolean allowNull) {
        return new ColumnSpecification(cellType, allowNull);
    }


    public static Builder builder(){
        return new Builder();
    }

    public static class Builder{

        private boolean ignoreColumn = false;
        private Class cellType;
        private boolean allowNull = true;
        private List<String> columns = new ArrayList<>();

        public Builder ignoreColumn(boolean ignoreColumn){
            this.ignoreColumn = ignoreColumn;
            return this;
        }

        public Builder cellType(Class cellType){
            this.cellType = cellType;
            return this;
        }

        public Builder allowNull(boolean allowNull){
            this.allowNull = allowNull;
            return this;
        }

        public Builder columns(List<String> columns){
            this.columns = columns;
            return this;
        }

        public ColumnSpecification build(){
            ColumnSpecification columnSpecification = new ColumnSpecification();
            columnSpecification.setCellType(cellType);
            columnSpecification.setColumns(columns);
            columnSpecification.setAllowNull(allowNull);
            columnSpecification.setIgnoreColumn(ignoreColumn);
            return columnSpecification;
        }
    }

}
