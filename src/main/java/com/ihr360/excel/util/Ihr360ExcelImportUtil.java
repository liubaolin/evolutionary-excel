package com.ihr360.excel.util;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.ihr360.excel.annotation.ExcelCell;
import com.ihr360.excel.config.ExcelConfigurer;
import com.ihr360.excel.config.Ihr360TemplateExcelConfiguration;
import com.ihr360.excel.context.Ihr360ImportExcelContext;
import com.ihr360.excel.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.exception.ExcelException;
import com.ihr360.excel.handler.Ihr360ExcelCellHandler;
import com.ihr360.excel.handler.Ihr360ExcelJavaBeanDataHandler;
import com.ihr360.excel.handler.Ihr360ExcelRowUtil;
import com.ihr360.excel.handler.Ihr360ExcelSpecificationHandler;
import com.ihr360.excel.handler.Ihr360ExcelValidatorHandler;
import com.ihr360.excel.logs.ExcelCommonLog;
import com.ihr360.excel.logs.ExcelLogItem;
import com.ihr360.excel.logs.ExcelLogType;
import com.ihr360.excel.logs.ExcelLogs;
import com.ihr360.excel.logs.ExcelRowLog;
import com.ihr360.excel.metaData.CellComment;
import com.ihr360.excel.metaData.ImportParams;
import com.ihr360.excel.specification.CommonSpecification;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Comparator;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static com.ihr360.excel.handler.Ihr360ExcelValidatorHandler.judgeHeader;


/**
 * The <code>Ihr360ExcelImportUtil</code> 与 {@link ExcelCell}搭配使用
 *
 * @author richey.liu
 * @version 1.0, Created at 2017-12-17
 */
public class Ihr360ExcelImportUtil {

    private static Logger logger = LoggerFactory.getLogger(Ihr360ExcelImportUtil.class);

    private Ihr360ExcelImportUtil() {
    }


    /**
     * 把Excel的数据封装成voList
     *
     * @param inputStream excel输入流
     * @return voList
     * @throws RuntimeException
     */
    @Deprecated
    public static <T> Collection<T> importExcel(ImportParams<T> importParams, InputStream inputStream
    ) {
        Ihr360ImportExcelContextHolder.initImportContext(importParams, inputStream);
        return importExcel();
    }


    public static Collection importExcel(ExcelConfigurer excelConfigurer, InputStream inputStream) {

        ImportParams importParams = excelConfigurer.getImportParam();

        if (excelConfigurer instanceof Ihr360TemplateExcelConfiguration) {
            Ihr360TemplateExcelConfiguration ihr360TemplateExcelConfiguration = (Ihr360TemplateExcelConfiguration) excelConfigurer;
            CommonSpecification commonSpecification = importParams.getCommonSpecification();
            commonSpecification.setTemplateHeaderRowNums(ihr360TemplateExcelConfiguration.getTemplateHeaderRowNums());
            commonSpecification.setTemplateDataBeginRowNum(ihr360TemplateExcelConfiguration.getTemplateDataBeginRowNum());
            commonSpecification.setTemplateHeaders(ihr360TemplateExcelConfiguration.getTemplateHeaders());
        } else {
            // todo
        }

        return importExcel(importParams, inputStream);
    }

    /**
     * 历史版本，不建议使用，可使用  {@link #importExcel(ExcelConfigurer excelConfigurer, InputStream inputStream)}
     *
     * @param <T>
     * @return
     */
    @Deprecated
    public static <T> Collection<T> importExcel() {

        Ihr360ImportExcelContext<T> ihr360ImportExcelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams<T> importParams = ihr360ImportExcelContext.getImportParams();
        ExcelLogs logs = ihr360ImportExcelContext.getLogs();
        Class<T> clazz = importParams.getImportType();

        Sheet sheet = ihr360ImportExcelContext.getCurrentSheet();
        if (sheet == null) {
            return new ArrayList<>();
        }

        List<T> resultList = new ArrayList<>();
        List<ExcelRowLog> rowLogList = new ArrayList<>();

        Iterator<Row> rowIterator = sheet.rowIterator();
        // 从excel读取的表头 Map<title,index>
        Map<String, Integer> headerTitleIndexMap = new LinkedHashMap<>();
        //最后的连续空行忽略
        List<ExcelRowLog> tempEmptyLogList = new ArrayList<>();

        CommonSpecification commonSpecification = importParams.getCommonSpecification();

        List<List<String>> headerJudgeList = Lists.newArrayList();
        if (commonSpecification != null) {
            headerJudgeList = commonSpecification.getHeaderColumnJudge();
            if (MapUtils.isNotEmpty(commonSpecification.getTemplateHeaders())) {
                headerTitleIndexMap = commonSpecification.getTemplateHeaders();
            }
        }

        Row headerRow = null;
        List<Integer> convertedHeaderRows = Lists.newArrayList();
        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            if (Ihr360ExcelRowUtil.ignorImportRow(row) && commonSpecification.getTemplateHeaders().size() != convertedHeaderRows.size()) {
                continue;
            }

            boolean isBlankRow = Ihr360ExcelRowUtil.checkBlankRow(row);
            boolean isTemplateHeaderRow = Ihr360ExcelRowUtil.isTemplateHeaderRow(row);
            List<ExcelLogItem> rowLogItems = new ArrayList<>();

            //处理表头
            if (isTemplateHeaderRow) {
                headerTitleIndexMap = getHeaderIndexTitleMap(headerTitleIndexMap, convertedHeaderRows, headerRow, row);
                continue;
            } else if (Ihr360ExcelRowUtil.isHeaderRow(headerRow, row)) {
                //隐藏行或空行
                if (Ihr360ExcelRowUtil.isHiddenOrBlanRow(row)) {
                    rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.FIRST_ROW_RULE, null));
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                    break;
                }
                headerTitleIndexMap = Ihr360ExcelRowUtil.convertRowToHeaderMap(row);
                headerRow = row;

                //表头判断
                boolean isHeader = judgeHeader(headerTitleIndexMap, headerJudgeList);

                if (!isHeader) {
                    headerRow = null;
                    rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.IGNORE_ROW, new String[]{row.getRowNum() + 1 + ""}));
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                }
                continue;
            }

            if (headerRow == null) {
                continue;
            }

            //设置headerRowNum
            if (commonSpecification != null && CollectionUtils.isNotEmpty(commonSpecification.getTemplateHeaderRowNums())) {
                ihr360ImportExcelContext.setHeaderRowNum(commonSpecification.getTemplateHeaderRowNums().stream().max(Comparator.naturalOrder()).get());
            } else {
                ihr360ImportExcelContext.setHeaderRowNum(headerRow.getRowNum());
            }


            //隐藏行
            if (row.getZeroHeight() && Ihr360ExcelRowUtil.ignorHiddenRows()) {
                rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.HIDDEN_ROW, new String[]{row.getRowNum() + 1 + ""}));
                rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                continue;
            }
            //根据规则忽略行
            List<List<String>> atLeastHeaders = commonSpecification == null ? null : commonSpecification.getAtLeastOneOrIgnoreRow();
            if (CollectionUtils.isNotEmpty(atLeastHeaders)) {
                Set<String> headerSet = headerTitleIndexMap.keySet();
                boolean contains = false;
                for (List<String> ailiasHeaders : atLeastHeaders) {
                    if (CollectionUtils.isEmpty(ailiasHeaders)) {
                        continue;
                    }
                    for (String ailiasHeader : ailiasHeaders) {
                        for (String header : headerSet) {
                            String headerOld = header;
                            if (!Ihr360ExcelValidatorHandler.headerEqueals(header, ailiasHeader)) {
                                continue;
                            }
                            Integer index = headerTitleIndexMap.get(headerOld);
                            Cell cell = row.getCell(index);
                            if (!Ihr360ExcelCellHandler.isNullOrBlankStringCell(cell)) {
                                contains = true;
                                break;
                            }
                        }
                    }
                    if (contains) {
                        break;
                    }
                }
                if (!contains) {
                    rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.IGNORE_ROW, new String[]{row.getRowNum() + 1 + ""}));
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                    continue;
                }
            }

            // 跳过空行,并记录日志
            if (isBlankRow) {
                rowLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.BLANK_ROW, new String[]{row.getRowNum() + 1 + ""}));
                tempEmptyLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                continue;
            } else {
                if (CollectionUtils.isNotEmpty(tempEmptyLogList)) {
                    rowLogList.addAll(tempEmptyLogList);
                    tempEmptyLogList.clear();
                }
            }

            //输出数据类型是Map时，简单将数据封装为Map<headerName,value>
            if (clazz == Map.class) {
                Map<String, Object> map = Ihr360ExcelRowUtil.handleExcelRowToMap(headerTitleIndexMap, row, rowLogItems, importParams.getColumnSpecifications());
                if (CollectionUtils.isEmpty(rowLogItems)) {
                    Ihr360ExcelSpecificationHandler.handleCommonSpecification(importParams, row, map);
                    resultList.add((T) map);
                } else {
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                }
            } else {

                T excelEntityVo = Ihr360ExcelJavaBeanDataHandler.handleImportExcelRowToJavabean(importParams, headerTitleIndexMap, rowLogItems, row);
                if (CollectionUtils.isEmpty(rowLogItems)) {
                    resultList.add(excelEntityVo);
                } else {
                    rowLogList.add(new ExcelRowLog(rowLogItems, row.getRowNum() + 1));
                }
            }
        }
        if (headerRow == null && MapUtils.isEmpty(headerTitleIndexMap) && CollectionUtils.isNotEmpty(headerJudgeList)) {
            ExcelCommonLog commonLog = new ExcelCommonLog();
            logs.setExcelCommonLog(commonLog);
            List<String> headers = headerJudgeList.stream().map(header -> header.get(0)).collect(Collectors.toList());

            List<ExcelLogItem> excelLogItems = Lists.newArrayList();
            ExcelLogItem excelLogItem = new ExcelLogItem(ExcelLogType.HEADER_REQUIRED, headers.toArray());
            excelLogItems.add(excelLogItem);
            commonLog.setExcelLogItems(excelLogItems);
            return resultList;
        }

        if (CollectionUtils.isEmpty(resultList) && CollectionUtils.isEmpty(rowLogList)) {
            ExcelCommonLog commonLog = new ExcelCommonLog();
            logs.setExcelCommonLog(commonLog);
            List<ExcelLogItem> excelLogItems = Lists.newArrayList();
            ExcelLogItem excelLogItem = new ExcelLogItem(ExcelLogType.EXCEL_COMMON_NO_DATA, null);
            excelLogItems.add(excelLogItem);
            commonLog.setExcelLogItems(excelLogItems);
            return resultList;
        }


        logs.setRowLogList(rowLogList);
        return resultList;
    }


    /**
     * 获取Excel第一行数据
     * 不忽略空行与隐藏行
     *
     * @return
     */
    public static Map<String, Integer> importGetFirstLineHeaderToMap() {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        Sheet sheet = excelContext.getCurrentSheet();
        if (sheet == null) {
            return Maps.newLinkedHashMap();
        }
        Iterator<Row> rowIterator = sheet.rowIterator();
        // 从excel读取的表头 Map<title,index>
        Map<String, Integer> fileHeaderIndexMap = Maps.newLinkedHashMap();

        if (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (row.getRowNum() == 0) {
                fileHeaderIndexMap = Ihr360ExcelRowUtil.convertRowToHeaderMap(row);
            }
        }
        return fileHeaderIndexMap;
    }

    /**
     * 导入获取真实数据所有条数
     *
     * @return
     */
    public static Integer importGetDataNum() {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        Sheet sheet = excelContext.getCurrentSheet();
        Integer dataNum = 0;

        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (!Ihr360ExcelRowUtil.checkBlankRow(row) && !row.getZeroHeight()) {
                dataNum++;
            }
        }
        return dataNum;
    }

    /**
     * 根据日志生成带注释的表格
     *
     * @param rowLogList
     * @param removeRight              删除正确数据
     * @param ihr360ImportExcelContext
     * @return
     */
    public static byte[] generateCommentFile(InputStream inputStream, List<ExcelRowLog> rowLogList, boolean removeRight, Ihr360ImportExcelContext ihr360ImportExcelContext) {

        if (CollectionUtils.isEmpty(rowLogList)) {
            throw new ExcelException("日志为空，无法生成注释文件！");
        }
        Map<Integer, List<ExcelLogItem>> rowLogMap = Maps.newHashMap();
        rowLogList.forEach(rowLog -> {
            rowLogMap.put(rowLog.getRowNum(), rowLog.getExcelLogItems());
        });
        Sheet sheet = null;
        try (Workbook workBook = WorkbookFactory.create(new BufferedInputStream(inputStream));) {
            sheet = workBook.getSheetAt(0);
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            Iterator<Row> rowIterator = sheet.rowIterator();
            Set<Row> toRevomeRows = Sets.newHashSet();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                List<ExcelLogItem> logItems = rowLogMap.get(row.getRowNum() + 1);

                if (CollectionUtils.isEmpty(logItems)) {
                    if (removeRight && ihr360ImportExcelContext != null && ihr360ImportExcelContext.getHeaderRowNum() < row.getRowNum()) {
                        toRevomeRows.add(row);
                    }
                    continue;
                }

                logItems = logItems.stream().filter(l -> l.getColNum() != null).collect(Collectors.toList());

                for (ExcelLogItem logItem : logItems) {
                    String commentStr = logItem.getMessage();
                    Integer columnIndex = logItem.getColNum();
                    Cell cell = row.getCell(columnIndex);
                    if (cell == null) {
                        cell = row.createCell(columnIndex);
                    }
                    CellComment cellComment = CellComment.createCellComment(new int[]{255, 125, 1023, 150, 0, 0, 2, 2}, null, commentStr, false);


                    Drawing patr = sheet.createDrawingPatriarch();
                    int[] params = cellComment.getAnchorParams();
                    ClientAnchor anchor = patr.createAnchor(params[0], params[1], params[2], params[3], params[4], params[5], params[6], params[7]);
                    Comment comment = patr.createCellComment(anchor);
                    if (StringUtils.isNotBlank(cellComment.getContentString())) {
                        try {
                            comment.setString(new HSSFRichTextString(cellComment.getContentString()));
                        } catch (IllegalArgumentException exception) {
                            comment.setString(new XSSFRichTextString(cellComment.getContentString()));
                        }
                    }
                    if (StringUtils.isNotBlank(cellComment.getAuthor())) {
                        comment.setAuthor(cellComment.getAuthor());
                    }
                    comment.setVisible(cellComment.isVisible());
                    cell.setCellComment(comment);
                }
            }
            if (CollectionUtils.isNotEmpty(toRevomeRows)) {
                for (Row row : toRevomeRows) {
                    sheet.getRow(row.getRowNum()).setZeroHeight(true);
                }
            }

            workBook.write(out);
            return out.toByteArray();

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        if (sheet == null) {
            throw new ExcelException("获取文件失败！");
        }
        return null;
    }


    private static Map<String, Integer> getHeaderIndexTitleMap(Map<String, Integer> headerTitleIndexMap,
                                                               final List<Integer> convertedHeaderRows, Row headerRow, Row row) {

        headerRow = row;

        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams<T> importParams = excelContext.getImportParams();
        CommonSpecification commonSpecification = importParams.getCommonSpecification();

        for (Integer headerRowNum : commonSpecification.getTemplateHeaderRowNums()) {
            if (convertedHeaderRows.contains(headerRowNum)) {
                continue;
            }
            headerRow = row;
            convertedHeaderRows.add(row.getRowNum());

            if (MapUtils.isEmpty(headerTitleIndexMap)) {
                headerTitleIndexMap = Ihr360ExcelRowUtil.convertRowToHeaderMap(headerRow);
            } else {
                Map<Integer, String> existIndexHeaderMap = new LinkedHashMap<>();
                headerTitleIndexMap.forEach((header, index) -> {
                    existIndexHeaderMap.put(index, header);
                });

                Map<String, Integer> newHeaderIndexMap = Ihr360ExcelRowUtil.convertRowToHeaderMap(headerRow);

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
