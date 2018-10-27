package com.ihr360.excel.commons.context;

import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.commons.logs.ExcelLogs;
import com.ihr360.excel.core.metaData.ImportParams;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.core.NamedThreadLocal;

import java.io.IOException;
import java.io.InputStream;

/**
 * @author richey
 */
public class Ihr360ImportExcelContextHolder {

    private static final ThreadLocal<Ihr360ImportExcelContext> importExcelContextHolder = new NamedThreadLocal("import excel context");

    private static final Logger logger = LoggerFactory.getLogger(Ihr360ImportExcelContextHolder.class);


    public Ihr360ImportExcelContextHolder() {
    }

    public static <T> void setExcelContext(Ihr360ImportExcelContext<T> excelContext) {
        if (excelContext == null) {
            importExcelContextHolder.remove();
        } else {
            importExcelContextHolder.set(excelContext);
        }
    }


    public static <T> Ihr360ImportExcelContext<T> getExcelContext() {
        Ihr360ImportExcelContext excelContext = importExcelContextHolder.get();
        if (excelContext == null) {
            excelContext = new Ihr360ImportExcelContext();
            importExcelContextHolder.set(excelContext);
        }
        return excelContext;
    }


    public static <T> void initImportContext(ImportParams<T> importParams, InputStream inputStream) {

        if (importParams == null) {
            throw new ExcelException("导入参数(importParams)不能为空,请检查");
        }
        if (inputStream == null) {
            throw new ExcelException("文件输出流不能为空");
        }

        Ihr360ImportExcelContext excelContext = getExcelContext();
        excelContext.setImportParams(importParams);
        excelContext.setInputStream(inputStream);
        excelContext.setLogs(new ExcelLogs());
    }

    public static void clean() {
        Ihr360ImportExcelContext excelContext = getExcelContext();
        InputStream inputStream = excelContext.getInputStream();
        if (inputStream != null) {
            try {
                inputStream.close();
            } catch (IOException e) {
                logger.error("输入流关闭失败", e);
            }
        }
        importExcelContextHolder.remove();
    }


}
