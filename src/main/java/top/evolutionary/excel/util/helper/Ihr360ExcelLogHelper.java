package top.evolutionary.excel.util.helper;

import top.evolutionary.excel.commons.context.Ihr360ImportExcelContext;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContextHolder;
import top.evolutionary.excel.commons.logs.ExcelLogItem;
import top.evolutionary.excel.commons.logs.ExcelLogType;
import top.evolutionary.excel.commons.logs.ExcelLogs;
import top.evolutionary.excel.commons.logs.ExcelRowLog;

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
