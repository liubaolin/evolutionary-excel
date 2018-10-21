package com.ihr360.excel.logs;

import java.util.List;

/**
 * The <code>ExcelRowLog</code>
 * <p>
 *     Excel行日志
 * </p>
 * @author richey.liu
 * @version 1.0, Created at 2017-12-17
 */
public class ExcelRowLog {

    private List<ExcelLogItem> excelLogItems;

    /**
     * 行号
     */
    private Integer rowNum;
    /**
     * 操作的对象
     */
    private Object object;


    /**
     * @return the rowNum
     */
    public Integer getRowNum() {
        return rowNum;
    }

    /**
     * @param rowNum the rowNum to set
     */
    public void setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
    }

    /**
     * @return the object
     */
    public Object getObject() {
        return object;
    }

    /**
     * @param object the object to set
     */
    public void setObject(Object object) {
        this.object = object;
    }

    public List<ExcelLogItem> getExcelLogItems() {
        return excelLogItems;
    }

    public void setExcelLogItems(List<ExcelLogItem> excelLogItems) {
        this.excelLogItems = excelLogItems;
    }


    public ExcelRowLog() {
    }

    /**
     * @param object
     * @param excelLogItems
     */
    public ExcelRowLog(Object object, List<ExcelLogItem> excelLogItems) {
        this.object = object;
        this.excelLogItems = excelLogItems;
    }

    /**
     * @param rowNum
     * @param object
     * @param excelLogItems
     */
    public ExcelRowLog(Object object, List<ExcelLogItem> excelLogItems, Integer rowNum) {
        this.object = object;
        this.rowNum = rowNum;
        this.excelLogItems = excelLogItems;
    }

    public ExcelRowLog(List<ExcelLogItem> excelLogItems, Integer rowNum) {
        this.excelLogItems = excelLogItems;
        this.rowNum = rowNum;
    }

}
