package top.evolutionary.excel.processor;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContext;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContextHolder;
import top.evolutionary.excel.commons.logs.ExcelCommonLog;
import top.evolutionary.excel.commons.logs.ExcelLogItem;
import top.evolutionary.excel.commons.logs.ExcelLogType;
import top.evolutionary.excel.commons.logs.ExcelLogs;
import top.evolutionary.excel.core.metaData.ImportParams;
import top.evolutionary.excel.commons.specification.CommonSpecification;
import top.evolutionary.excel.util.helper.Ihr360ExcelLogHelper;
import top.evolutionary.excel.util.helper.Ihr360ExcelRowHelper;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static top.evolutionary.excel.util.helper.Ihr360ExcelValidatorHelper.judgeHeader;

/**
 * @author richey
 */
public class Ihr360ExcelHeaderProcessor<T> extends AbstractIhr360ImportExcelProcessor {

    public Ihr360ExcelHeaderProcessor() {
        super(Integer.MIN_VALUE);
    }

    public Ihr360ExcelHeaderProcessor(int order) {
        super(order);
    }

    @Override
    public void doProcess() {

        Ihr360ImportExcelContext<T> ihr360ImportExcelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams<T> importParams = ihr360ImportExcelContext.getImportParams();
        CommonSpecification commonSpecification = importParams.getCommonSpecification();


        Sheet sheet = ihr360ImportExcelContext.getCurrentSheet();
        if (sheet == null) {
            setResult(Collections.emptyList());
            setBack(true);
            return;
        }

        //处理表头行
        processHeaderRow();
        Map<String, Integer> headerTitleIndexMap = ihr360ImportExcelContext.getHeaderTitleIndexMap();
        List<List<String>> headerJudgeList = commonSpecification == null ? Lists.newArrayList() : commonSpecification.getHeaderColumnJudge();

        ExcelLogs logs = ihr360ImportExcelContext.getLogs();
        ExcelCommonLog commonLog = logs.getExcelCommonLog();
        if (MapUtils.isEmpty(headerTitleIndexMap) && CollectionUtils.isNotEmpty(headerJudgeList)) {
            List<String> headers = headerJudgeList.stream().map(header -> header.get(0)).collect(Collectors.toList());
            List<ExcelLogItem> excelLogItems = Lists.newArrayList();
            ExcelLogItem excelLogItem = new ExcelLogItem(ExcelLogType.HEADER_REQUIRED, headers.toArray());
            excelLogItems.add(excelLogItem);
            commonLog.setExcelLogItems(excelLogItems);
            super.setBack(true);
            super.setResult(Collections.EMPTY_LIST);
        }
    }

    private void processHeaderRow() {

        Row headerRow = null;
        List<Integer> convertedHeaderRows = Lists.newArrayList();
        Ihr360ImportExcelContext<org.apache.poi.ss.formula.functions.T> excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        CommonSpecification commonSpecification = excelContext.getImportParams().getCommonSpecification();
        List<List<String>> headerJudgeList = commonSpecification == null ? Lists.newArrayList() : commonSpecification.getHeaderColumnJudge();

        Sheet sheet = excelContext.getCurrentSheet();
        Iterator<Row> rowIterator = sheet.rowIterator();
        Map<String, Integer> headerTitleIndexMap = null;
        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();
            if (Ihr360ExcelRowHelper.ignorImportRow(row) && commonSpecification.getTemplateHeaderIndexTitleMap().size() != convertedHeaderRows.size()) {
                continue;
            }
            boolean isTemplateHeaderRow = Ihr360ExcelRowHelper.isTemplateHeaderRow(row);
            //处理表头
            if (isTemplateHeaderRow) {
                headerTitleIndexMap = getHeaderIndexTitleMap(convertedHeaderRows, headerRow, row);
                continue;
            } else if (Ihr360ExcelRowHelper.isHeaderRow(headerRow, row)) {
                //隐藏行或空行
                if (Ihr360ExcelRowHelper.isHiddenOrBlanRow(row)) {
                    Ihr360ExcelLogHelper.addToRowLogList(ExcelLogType.FIRST_ROW_RULE, null, row.getRowNum());
                    break;
                }
                headerTitleIndexMap = Ihr360ExcelRowHelper.convertRowToHeaderTitleIndexMap(row);
                headerRow = row;

                //表头判断
                boolean isHeader = judgeHeader(headerTitleIndexMap, headerJudgeList);

                if (!isHeader) {
                    headerRow = null;
                    Ihr360ExcelLogHelper.addToRowLogList(ExcelLogType.IGNORE_ROW, new String[]{row.getRowNum() + 1 + ""}, row.getRowNum());
                }
                continue;
            }

            if (headerRow != null) {
                excelContext.setHeaderRowNum(headerRow.getRowNum());
                break;
            }
        }

        excelContext.setHeaderTitleIndexMap(headerTitleIndexMap);
    }

    private Map<String, Integer> getHeaderIndexTitleMap(final List<Integer> convertedHeaderRows, Row headerRow, Row row) {

        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        CommonSpecification commonSpecification = excelContext.getImportParams().getCommonSpecification();
        Map<String, Integer> templateHeaderIndexTitleMap = commonSpecification.getTemplateHeaderIndexTitleMap();
        Map<String, Integer> headerTitleIndexMap = Maps.newHashMap();

        for (Integer headerRowNum : commonSpecification.getTemplateHeaderRowNums()) {
            if (convertedHeaderRows.contains(headerRowNum)) {
                continue;
            }
            headerRow = row;
            convertedHeaderRows.add(row.getRowNum());

            if (MapUtils.isEmpty(templateHeaderIndexTitleMap)) {
                headerTitleIndexMap = Ihr360ExcelRowHelper.convertRowToHeaderTitleIndexMap(headerRow);
            } else {
                Map<Integer, String> existIndexHeaderMap = new LinkedHashMap<>();
                templateHeaderIndexTitleMap.forEach((header, index) -> {
                    existIndexHeaderMap.put(index, header);
                });

                Map<String, Integer> newHeaderIndexMap = Ihr360ExcelRowHelper.convertRowToHeaderTitleIndexMap(headerRow);

                if (MapUtils.isNotEmpty(newHeaderIndexMap)) {
                    newHeaderIndexMap.forEach((header, index) -> {
                        existIndexHeaderMap.put(index, header);

                    });
                    Map<String, Integer> tmpHeaderIndexMap = new LinkedHashMap<>();
                    existIndexHeaderMap.forEach((index, header) -> {
                        tmpHeaderIndexMap.put(header, index);
                    });
                    headerTitleIndexMap = tmpHeaderIndexMap;
                }
            }
            break;
        }

        return headerTitleIndexMap;
    }

}
