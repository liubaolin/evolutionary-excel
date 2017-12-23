package org.joyful4j.modules.excel.entity;

import org.joyful4j.modules.excel.annotation.ExcelCell;

import java.util.Map;

/**
 * 支持弹性字段导入的Excel
 * @author richey
 */
public abstract class AbstractFlexbleFieldExcel {

    @ExcelCell(flexibleField = true)
    protected Map<String, Object> flexbleFields;

    public Map<String, Object> getFlexbleFields() {
        return flexbleFields;
    }

    public void setFlexbleFields(Map<String, Object> flexbleFields) {
        this.flexbleFields = flexbleFields;
    }


}
