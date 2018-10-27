package com.ihr360.excel.processor;

import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.commons.logs.ExcelLogItem;
import com.ihr360.excel.commons.logs.ExcelLogType;
import com.ihr360.excel.commons.logs.ExcelLogs;
import com.ihr360.excel.commons.logs.ExcelRowLog;
import com.ihr360.excel.core.metaData.ImportParams;
import com.ihr360.excel.commons.specification.CommonSpecification;
import com.ihr360.excel.util.helper.Ihr360ExcelCellHelper;
import com.ihr360.excel.util.helper.Ihr360ExcelJavaBeanDataHelper;
import com.ihr360.excel.util.helper.Ihr360ExcelLogHelper;
import com.ihr360.excel.util.helper.Ihr360ExcelRowHelper;
import com.ihr360.excel.util.helper.Ihr360ExcelSpecificationHelper;
import com.ihr360.excel.util.helper.Ihr360ExcelValidatorHelper;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @author richey
 */
public class Ihr360ExcelRowDataProcessor<T> extends AbstractIhr360ImportExcelProcessor {


    public Ihr360ExcelRowDataProcessor(int order) {
        super(order);
    }

    @Override
    public void doProcess() {
        Ihr360ImportExcelContext<T> ihr360ImportExcelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams<T> importParams = ihr360ImportExcelContext.getImportParams();

        List<ExcelRowLog> tempEmptyRowLogList = new ArrayList<>();
        Sheet sheet = ihr360ImportExcelContext.getCurrentSheet();
        Iterator<Row> rowIterator = sheet.rowIterator();
        CommonSpecification commonSpecification = importParams.getCommonSpecification();
        Map<String, Integer> headerTitleIndexMap = ihr360ImportExcelContext.getHeaderTitleIndexMap();

        List<T> resultList = new ArrayList<>();
        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();
            if (row.getRowNum() <= ihr360ImportExcelContext.getHeaderRowNum()) {
                continue;
            }

            //设置headerRowNum
            if (commonSpecification != null && CollectionUtils.isNotEmpty(commonSpecification.getTemplateHeaderRowNums())) {
                ihr360ImportExcelContext.setHeaderRowNum(commonSpecification.getTemplateHeaderRowNums().stream().max(Comparator.naturalOrder()).get());
            } else {
                ihr360ImportExcelContext.setHeaderRowNum(ihr360ImportExcelContext.getHeaderRowNum());
            }

            //隐藏行
            if (row.getZeroHeight() && Ihr360ExcelRowHelper.ignorHiddenRows()) {
                Ihr360ExcelLogHelper.addToRowLogList(ExcelLogType.HIDDEN_ROW, new String[]{row.getRowNum() + 1 + ""}, row.getRowNum());
                continue;
            }
            //根据规则忽略行
            List<List<String>> atLeastOneHeaderTitles = commonSpecification == null ? null : commonSpecification.getAtLeastOneOrIgnoreRow();
            if (CollectionUtils.isNotEmpty(atLeastOneHeaderTitles)) {
                Set<String> headerTitleSet = headerTitleIndexMap.keySet();
                boolean contains = false;
                for (List<String> ailiasHeaders : atLeastOneHeaderTitles) {
                    if (CollectionUtils.isEmpty(ailiasHeaders)) {
                        continue;
                    }
                    for (String ailiasHeader : ailiasHeaders) {
                        for (String header : headerTitleSet) {
                            if (!Ihr360ExcelValidatorHelper.headerEqueals(header, ailiasHeader)) {
                                continue;
                            }
                            Integer index = headerTitleIndexMap.get(header);
                            Cell cell = row.getCell(index);
                            if (Ihr360ExcelCellHelper.isNullOrBlankStringCell(cell)) {
                                continue;
                            }
                            contains = true;
                            break;
                        }
                    }
                    if (contains) {
                        break;
                    }
                }
                if (!contains) {
                    Ihr360ExcelLogHelper.addToRowLogList(ExcelLogType.IGNORE_ROW, new String[]{row.getRowNum() + 1 + ""}, row.getRowNum());
                    continue;
                }
            }

            // 跳过空行,并记录日志,忽略最后的连续空行忽略
            if (Ihr360ExcelRowHelper.checkBlankRow(row)) {
                List<ExcelLogItem> rowLogItems = new ArrayList<>();
                rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.BLANK_ROW, new String[]{row.getRowNum() + 1 + ""}));
                tempEmptyRowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                continue;
            } else if (CollectionUtils.isNotEmpty(tempEmptyRowLogList)) {
                ExcelLogs logs = ihr360ImportExcelContext.getLogs();
                List<ExcelRowLog> rowLogList = logs.getRowLogList();
                rowLogList.addAll(tempEmptyRowLogList);
                tempEmptyRowLogList.clear();
            }

            //输出数据类型是Map时，简单将数据封装为Map<headerName,value>
            List<ExcelLogItem> rowLogItems = new ArrayList<>();

            Class<T> clazz = importParams.getImportType();
            if (clazz == Map.class) {

                Map<String, Object> map = Ihr360ExcelRowHelper.handleExcelRowToMap(headerTitleIndexMap, row, rowLogItems);
                if (CollectionUtils.isEmpty(rowLogItems)) {
                    Ihr360ExcelSpecificationHelper.handleCommonSpecification(row, map);
                    resultList.add((T) map);
                } else {
                    ExcelLogs logs = ihr360ImportExcelContext.getLogs();
                    List<ExcelRowLog> rowLogList = logs.getRowLogList();
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                }
            } else {
                T excelEntityVo = Ihr360ExcelJavaBeanDataHelper.handleImportExcelRowToJavabean(headerTitleIndexMap, rowLogItems, row);
                if (CollectionUtils.isEmpty(rowLogItems)) {
                    resultList.add(excelEntityVo);
                } else {
                    ExcelLogs logs = ihr360ImportExcelContext.getLogs();
                    List<ExcelRowLog> rowLogList = logs.getRowLogList();
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                }
            }
        }
        setResult(resultList);
    }

}
