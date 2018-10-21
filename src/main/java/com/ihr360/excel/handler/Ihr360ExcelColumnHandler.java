package com.ihr360.excel.handler;

import com.ihr360.excel.context.Ihr360ExportExcelContextHolder;
import com.ihr360.excel.context.Ihr360ImportExcelContext;
import com.ihr360.excel.metaData.ExportParams;
import com.ihr360.excel.specification.ExportCommonSpecification;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;

/**
 * 列处理类
 *
 * @author richey
 */
public class Ihr360ExcelColumnHandler {

    public static void setColumnHidden() {
        Ihr360ImportExcelContext ihr360ImportExcelContext = Ihr360ExportExcelContextHolder.getExcelContext();
        Sheet sheet = ihr360ImportExcelContext.getCurrentSheet();
        ExportParams<T> exportParams = ihr360ImportExcelContext.getExportParams();
        ExportCommonSpecification exportCommonSpecification = exportParams.getExportCommonSpecification();
        Map<String, String> headerMap = exportParams.getHeaderMap();

        if (sheet == null || exportCommonSpecification == null || MapUtils.isEmpty(headerMap)) {
            return;
        }
        List<String> hiddenColumns = exportCommonSpecification.getHiddenColumns();
        if (CollectionUtils.isEmpty(hiddenColumns)) {
            return;
        }

        int columnIndex = 0;
        for (Map.Entry<String, String> stringStringEntry : headerMap.entrySet()) {
            String columnKey = stringStringEntry.getKey();
            if (hiddenColumns.contains(columnKey)) {
                sheet.setColumnHidden(columnIndex++, true);
                continue;
            }
            columnIndex++;
        }
    }


}
