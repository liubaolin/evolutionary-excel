package com.ihr360.excel.util;

import com.ihr360.excel.context.Ihr360ImportExcelContext;
import com.ihr360.excel.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.logs.ExcelLogItem;
import com.ihr360.excel.logs.ExcelLogType;
import com.ihr360.excel.logs.ExcelLogs;
import com.ihr360.excel.logs.ExcelRowLog;

import java.util.ArrayList;
import java.util.List;

/**
 * @author richey
 */
public class Ihr360ExcelLogUtil {

    public static void addToRowLogList(ExcelLogType excelLogType, Object[] args, int rowNum) {
        List<ExcelLogItem> rowLogItems = new ArrayList<>();
        rowLogItems.add(ExcelLogItem.createExcelItem(excelLogType, args));
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ExcelLogs logs = excelContext.getLogs();
        List<ExcelRowLog> rowLogList = logs.getRowLogList();
        rowLogList.add(new ExcelRowLog(rowLogItems, rowNum + 1));
    }

}
