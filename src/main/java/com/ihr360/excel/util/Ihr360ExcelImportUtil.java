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
import com.ihr360.excel.handler.Ihr360ExcelHandlerManager;
import com.ihr360.excel.handler.Ihr360ExcelRowUtil;
import com.ihr360.excel.handler.Ihr360ImportExcelProcessor;
import com.ihr360.excel.logs.ExcelCommonLog;
import com.ihr360.excel.logs.ExcelLogItem;
import com.ihr360.excel.logs.ExcelLogType;
import com.ihr360.excel.logs.ExcelLogs;
import com.ihr360.excel.logs.ExcelRowLog;
import com.ihr360.excel.metaData.CellComment;
import com.ihr360.excel.metaData.ImportParams;
import com.ihr360.excel.specification.CommonSpecification;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;


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
            commonSpecification.setTemplateHeaderIndexTitleMap(ihr360TemplateExcelConfiguration.getTemplateHeaders());
        } else {
            // todo
        }

        return importExcel(importParams, inputStream);
    }

    private static <T> Collection<T> importExcel() {

        Ihr360ExcelHandlerManager<T> processorManager = new Ihr360ExcelHandlerManager();

        Collection<T> resultList = Lists.newArrayList();
        for (Ihr360ImportExcelProcessor<T> processor : processorManager.getProcessors()) {
            processor.doProcess();
            resultList = processor.getResult();
            if (processor.isBack()) {
                return resultList;
            }
        }
        
        Ihr360ImportExcelContext<T> ihr360ImportExcelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ExcelLogs logs = ihr360ImportExcelContext.getLogs();
        List<ExcelRowLog> rowLogList = logs.getRowLogList();
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


}
