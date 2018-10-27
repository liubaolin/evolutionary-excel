package com.ihr360.excel.commons.context;

import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.core.metaData.ExportParams;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.NamedThreadLocal;

import java.io.IOException;
import java.io.OutputStream;

/**
 * @author richey
 */
public class Ihr360ExportExcelContextHolder {

    private static final ThreadLocal<Ihr360ExportExcelContext> exportExcelContextHolder = new NamedThreadLocal("export excel context");

    private static final Logger logger = LoggerFactory.getLogger(Ihr360ExportExcelContextHolder.class);


    public Ihr360ExportExcelContextHolder() {
    }

    public static <T> void setExcelContext(Ihr360ExportExcelContext<T> excelContext) {
        if (excelContext == null) {
            exportExcelContextHolder.remove();
        } else {
            exportExcelContextHolder.set(excelContext);
        }
    }


    public static <T> Ihr360ExportExcelContext<T> getExcelContext() {
        Ihr360ExportExcelContext excelContext = exportExcelContextHolder.get();
        if (excelContext == null) {
            excelContext = new Ihr360ExportExcelContext();
            exportExcelContextHolder.set(excelContext);
        }
        return excelContext;
    }

    private static <T> void initExportHSSFWorkbookInstance() {
        Ihr360ExportExcelContext<T> excelContext = getExcelContext();
        Workbook workbook = excelContext.getCurrentWorkbook();
        if (workbook == null) {
            workbook = new HSSFWorkbook();
            excelContext.setCurrentWorkbook(workbook);
        }
        ExportParams<T> exportParams = excelContext.getExportParams();
        if (exportParams == null) {
            throw new ExcelException("导出参数(exportParams)不能为空,请检查");
        }
        if (CollectionUtils.isNotEmpty(exportParams.getSheets())) {
            //多Sheet导出目前需要从exportparam中传入sheets
            return;
        } else if (excelContext.getCurrentSheet() == null) {
            excelContext.setCurrentSheet(workbook.createSheet());
        }
    }


    public static <T> void initExportContext(boolean exportHssf, ExportParams<T> exportParams, OutputStream out) {
        clean();
        if (exportParams == null) {
            throw new ExcelException("导出参数(exportParams)不能为空,请检查");
        }
        if (out == null) {
            throw new ExcelException("导出操作输出流不能为空");
        }

        if (exportHssf) {
            Ihr360ExportExcelContext excelContext = getExcelContext();
            excelContext.setExportParams(exportParams);
            excelContext.setOutputStream(out);
            initExportHSSFWorkbookInstance();

        } else {
            throw new ExcelException("Not implement XSSFWorkbook Export  yet");
        }
    }

    public static void clean() {
        Ihr360ExportExcelContext excelContext = getExcelContext();
        OutputStream outputStream = excelContext.getOutputStream();
        if (outputStream != null) {
            try {
                outputStream.close();
            } catch (IOException e) {
                logger.error("导出操作输出流关闭失败", e);
            }
        }
        exportExcelContextHolder.remove();
    }


}
