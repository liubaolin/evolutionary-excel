package com.ihr360.excel.commons.logs;

import org.apache.commons.collections.CollectionUtils;

import java.util.ArrayList;
import java.util.List;

public class ExcelCommonLog {

    private List<ExcelLogItem> excelLogItems = new ArrayList<>();

    public ExcelCommonLog() {
    }

    public ExcelCommonLog(List<ExcelLogItem> excelLogItems) {
        this.excelLogItems = excelLogItems;
    }

    public List<ExcelLogItem> getExcelLogItems() {
        return excelLogItems;
    }

    public void setExcelLogItems(List<ExcelLogItem> excelLogItems) {
        this.excelLogItems = excelLogItems;
    }

    public boolean hasLogs() {
        return CollectionUtils.isNotEmpty(excelLogItems);
    }
}
