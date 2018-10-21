package com.ihr360.excel.context;

import com.google.common.collect.Maps;
import com.ihr360.excel.annotation.ExcelConfig;
import com.ihr360.excel.exception.ExcelException;
import com.ihr360.excel.logs.ExcelCommonLog;
import com.ihr360.excel.logs.ExcelLogItem;
import com.ihr360.excel.logs.ExcelLogType;
import com.ihr360.excel.logs.ExcelLogs;
import com.ihr360.excel.metaData.ImportParams;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * @author richey
 */
public class Ihr360ImportExcelContext<T> implements Serializable {

    private static final long serialVersionUID = 7716606597402878349L;

    private static final Logger logger = LoggerFactory.getLogger(Ihr360ExportExcelContextHolder.class);

    /**
     * 表头行号
     */
    private int headerRowNum;

    private ExcelLogs logs = new ExcelLogs();

    private ImportParams<T> importParams;

    private InputStream inputStream;

    private Workbook currentWorkbook;

    private Map<String, Sheet> sheets = Maps.newHashMap();

    private String sheetCursor = "0";

    public int getHeaderRowNum() {
        return headerRowNum;
    }

    public void setHeaderRowNum(int headerRowNum) {
        this.headerRowNum = headerRowNum;
    }

    public ExcelLogs getLogs() {
        return logs;
    }

    public void setLogs(ExcelLogs logs) {
        this.logs = logs;
    }

    public ImportParams<T> getImportParams() {
        if (importParams == null) {
            throw new ExcelException("导入参数不能为空");
        }
        return importParams;
    }

    public void setImportParams(ImportParams<T> importParams) {
        this.importParams = importParams;
    }

    public InputStream getInputStream() {
        return inputStream;
    }

    public void setInputStream(InputStream inputStream) {
        this.inputStream = new BufferedInputStream(inputStream);
    }

    public ExcelConfig getExcelConfig() {
        ImportParams importParams = getImportParams();
        Class<T> clazz = importParams.getImportType();
        ExcelConfig excelConfig = clazz.getAnnotation(ExcelConfig.class);
        return excelConfig;
    }


    public Sheet getCurrentSheet() {
        Sheet sheet = null;
        if (sheets == null || sheets.size() - 1 < Integer.parseInt(sheetCursor)) {
            Workbook workbook = getWorkbook();
            sheet = workbook.getSheet(sheetCursor);
            sheets.put(sheetCursor, sheet);
        }
        return sheet;
    }

    public Workbook getWorkbook() {
        if (currentWorkbook != null) {
            return currentWorkbook;
        }

        ExcelCommonLog commonLog = logs.getExcelCommonLog();
        List<ExcelLogItem> excelLogItems = commonLog.getExcelLogItems();
        try {
            //支持.xls和.xlsx
            currentWorkbook = WorkbookFactory.create(inputStream);
        } catch (EncryptedDocumentException e) {
            logger.info(e.getMessage());
            excelLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.EXCEL_COMMON_ENCRYPTED, null));
        } catch (InvalidFormatException e) {
            logger.info(e.getMessage());
            excelLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.EXCEL_COMMON_FORMAT_ENCRYPTED, null));
        } catch (Exception e) {
            logger.error("load excel file error", e);
            excelLogItems.add(ExcelLogItem.createExcelItem(ExcelLogType.EXCEL_COMMON_COMMON, null));
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                logger.error(e.toString());
            }
        }

        return currentWorkbook;
    }
}
