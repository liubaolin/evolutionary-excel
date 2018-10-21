package com.ihr360.excel.logs;

import org.apache.commons.collections.CollectionUtils;

import java.util.ArrayList;
import java.util.List;

/**
 * The <code>ExcelLogs</code>
 * <p>
 * <p>
 * </p>
 *
 * @author richey.liu
 * @version 1.0, Created at 2017-12-17
 */
public class ExcelLogs {

    /**
     * 行日志
     */
    private List<ExcelRowLog> rowLogList = new ArrayList<>();

    /**
     * 普通日志
     */
    private ExcelCommonLog excelCommonLog = new ExcelCommonLog();

    /**
     * @return
     */
    public boolean hasRowLogList() {
        return CollectionUtils.isNotEmpty(this.rowLogList);
    }

    public boolean hasExcelLogs() {
        return CollectionUtils.isNotEmpty(excelCommonLog.getExcelLogItems());
    }


    /**
     * @return the rowLogList
     */
    public List<ExcelRowLog> getRowLogList() {
        return rowLogList;
    }


    /**
     * @param rowLogList the rowLogList to set
     */
    public void setRowLogList(List<ExcelRowLog> rowLogList) {
        this.rowLogList = rowLogList;
    }

    public ExcelCommonLog getExcelCommonLog() {
        return excelCommonLog;
    }

    public void setExcelCommonLog(ExcelCommonLog excelCommonLog) {
        this.excelCommonLog = excelCommonLog;
    }
}
