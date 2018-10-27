package com.ihr360.excel.util.helper;

import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.commons.logs.ExcelLogItem;
import com.ihr360.excel.commons.logs.ExcelLogType;
import com.ihr360.excel.commons.logs.ExcelLogs;
import com.ihr360.excel.commons.logs.ExcelRowLog;

import java.util.ArrayList;
import java.util.List;

/**
 * @author richey
 */
public class Ihr360ExcelLogHelper {

    public static void addToRowLogList(ExcelLogType excelLogType, Object[] args, int rowNum) {
        List<ExcelLogItem> rowLogItems = new ArrayList<>();
        rowLogItems.add(ExcelLogItem.createExcelItem(excelLogType, args));
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ExcelLogs logs = excelContext.getLogs();
        List<ExcelRowLog> rowLogList = logs.getRowLogList();
        rowLogList.add(new ExcelRowLog(rowLogItems, rowNum + 1));

    }

}
