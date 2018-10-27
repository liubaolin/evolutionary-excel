package com.ihr360.excel.commons.specification;

import java.util.List;

/**
 * Ｅｘｃｅｌ导出描述
 *
 * @author richey
 */
public class ExportCommonSpecification {

    /**
     * 隐藏列
     */
    private List<String> hiddenColumns;

    public List<String> getHiddenColumns() {
        return hiddenColumns;
    }

    public void setHiddenColumns(List<String> hiddenColumns) {
        this.hiddenColumns = hiddenColumns;
    }
}
