package com.ihr360.excel.util.helper;

import com.ihr360.excel.core.cellstyle.ExcelCellStyle;
import com.ihr360.excel.core.cellstyle.ExcelCellStyleFactory;
import com.ihr360.excel.core.cellstyle.Ihr360CellStyle;
import com.ihr360.excel.core.metaData.CellComment;
import com.ihr360.excel.core.metaData.ExportParams;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author richey
 */
public class Ihr360ExcelCommonExcelHelper {

    private static final ThreadLocal<DataFormatter> dataFormatterThreadLocal = new ThreadLocal<>();


    public static DataFormatter getDataFormatter() {
        DataFormatter formatter = dataFormatterThreadLocal.get();
        if (formatter == null) {
            formatter = new DataFormatter();
            dataFormatterThreadLocal.set(formatter);
        }
        return formatter;
    }

    public static Map<Integer, String> handleExportHeader(Workbook workbook, Row row, Map<String, Font> fontMap, ExportParams exportParams, int startIndex, Sheet sheet, final Map<Integer, Class> dateTypeIndexMap) {
        Map<Integer, String> datePatternIndexMap = new HashMap<>();
        Map<String, String> datePatternMap = exportParams.getDatePatternMap();
        Map<String, Class> dateTypeMap = exportParams.getDataTypeMap();

        Ihr360CellStyle defaultHeaderCellStyle = ExcelCellStyleFactory.createDefaultHeaderCellStyle();
        Map<String, CellStyle> excelCellStyleMap = new HashMap<>();
        Map<String, String> headers = exportParams.getHeaderMap();

        if (MapUtils.isNotEmpty(exportParams.getMergedHeaderMap())) {
            headers = exportParams.getMergedHeaderMap();
        }

        List<String> headerKeys = new ArrayList<>(headers.keySet());
        Map<String, ExcelCellStyle> headerStyleMap = exportParams.getHeaderStyleMap();

        Map<String, CellComment> headerCommentMap = exportParams.getHeaderCommentMap();

        for (int i = 0; i < headerKeys.size(); i++) {
            int index = i + startIndex;
            Cell cell = row.createCell(index);
            String headerKey = headerKeys.get(i);
            String headerName = headers.get(headerKey);

            HSSFRichTextString text = new HSSFRichTextString(headerName);
            cell.setCellValue(text);

            if (MapUtils.isNotEmpty(datePatternMap)) {
                String pattern = datePatternMap.get(headerKey);
                if (StringUtils.isNotBlank(pattern)) {
                    datePatternIndexMap.put(index, pattern);
                }
            }
            if (MapUtils.isNotEmpty(dateTypeMap)) {
                Class clazz = dateTypeMap.get(headerKey);
                if (clazz != null) {
                    dateTypeIndexMap.put(index, clazz);
                }
            }

            Ihr360ExcelCellCommentHelper.setHeaderComment(sheet, headerCommentMap, cell, headerKey);

            CellStyle poiCellStyle = Ihr360ExcelCellStyleHelper.getPoiCellStyle(workbook, fontMap, defaultHeaderCellStyle, excelCellStyleMap, headerStyleMap, headerKey);
            cell.setCellStyle(poiCellStyle);
        }
        return datePatternIndexMap;
    }


}
