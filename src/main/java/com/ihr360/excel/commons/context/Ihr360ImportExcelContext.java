package com.ihr360.excel.commons.context;

import com.google.common.collect.Maps;
import com.ihr360.excel.core.annotation.ExcelConfig;
import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.commons.logs.ExcelCommonLog;
import com.ihr360.excel.commons.logs.ExcelLogItem;
import com.ihr360.excel.commons.logs.ExcelLogType;
import com.ihr360.excel.commons.logs.ExcelLogs;
import com.ihr360.excel.core.metaData.ImportParams;
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

    private Map<Integer, Sheet> sheets = Maps.newHashMap();

    private Integer sheetCursor = 0;

    private  Map<String, Integer> headerTitleIndexMap;

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
        Sheet sheet;
        if (sheets == null || sheets.size() - 1 < sheetCursor) {
            Workbook workbook = getWorkbook();
            sheet = workbook.getSheetAt(sheetCursor);
            sheets.put(sheetCursor, sheet);
        }else{
            return sheets.get(sheetCursor);
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

    public Workbook getCurrentWorkbook() {
        return currentWorkbook;
    }

    public void setCurrentWorkbook(Workbook currentWorkbook) {
        this.currentWorkbook = currentWorkbook;
    }

    public Map<Integer, Sheet> getSheets() {
        return sheets;
    }

    public void setSheets(Map<Integer, Sheet> sheets) {
        this.sheets = sheets;
    }

    public Integer getSheetCursor() {
        return sheetCursor;
    }

    public void setSheetCursor(Integer sheetCursor) {
        this.sheetCursor = sheetCursor;
    }

    public Map<String, Integer> getHeaderTitleIndexMap() {
        return headerTitleIndexMap;
    }

    public void setHeaderTitleIndexMap(Map<String, Integer> headerTitleIndexMap) {
        this.headerTitleIndexMap = headerTitleIndexMap;
    }
}
