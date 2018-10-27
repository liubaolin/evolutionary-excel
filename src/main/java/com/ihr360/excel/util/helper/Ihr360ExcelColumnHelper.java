package com.ihr360.excel.util.helper;

import com.ihr360.excel.commons.context.Ihr360ExportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ExportExcelContextHolder;
import com.ihr360.excel.core.metaData.ExportParams;
import com.ihr360.excel.commons.specification.ExportCommonSpecification;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;
import java.util.Map;

/**
 * 列处理类
 *
 * @author richey
 */
public class Ihr360ExcelColumnHelper {

    public static void setColumnHidden() {
        Ihr360ExportExcelContext excelContext = Ihr360ExportExcelContextHolder.getExcelContext();
        Sheet sheet = excelContext.getCurrentSheet();
        ExportParams exportParams = excelContext.getExportParams();
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
