package com.ihr360.excel.util.helper;

import com.ihr360.excel.commons.context.Ihr360ExportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ExportExcelContextHolder;
import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.commons.specification.MergedRegionSpecification;
import com.ihr360.excel.config.ExcelDefaultConfig;
import com.ihr360.excel.core.annotation.ExcelCell;
import com.ihr360.excel.core.cellstyle.ExcelCellStyle;
import com.ihr360.excel.core.cellstyle.ExcelCellStyleFactory;
import com.ihr360.excel.core.cellstyle.Ihr360CellStyle;
import com.ihr360.excel.core.metaData.ExportHeaderParams;
import com.ihr360.excel.core.metaData.ExportParams;
import com.ihr360.excel.core.metaData.MergedExportData;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.collections.map.HashedMap;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static com.ihr360.excel.util.helper.Ihr360ExcelCellStyleHelper.getPoiCellStyle;
import static com.ihr360.excel.util.helper.Ihr360ExcelCellStyleHelper.setSheetAutoAddWidthColumn;

/**
 * Sheet处理类
 *
 * @author richey
 */
public class Ihr360ExcelSheetHelper {

    private static Logger logger = LoggerFactory.getLogger(Ihr360ExcelSheetHelper.class);




    /**
     * 每个sheet的写入
     *
     */
    public static <T> void write2Sheet() {

        //设置隐藏列
        Ihr360ExcelColumnHelper.setColumnHidden();

        Ihr360ExportExcelContext<T> excelContext = Ihr360ExportExcelContextHolder.getExcelContext();
        Sheet sheet = excelContext.getCurrentSheet();
        Workbook workbook = excelContext.getCurrentWorkbook();
        ExportParams<T> exportParams = excelContext.getExportParams();
        Map<String, String> headers = exportParams.getHeaderMap();
        Collection<T> datas = exportParams.getRowDatas();

        // 标题行转中文 headers是有序的map，产生的keySet也是有序的
        List<String> headerKeys = new ArrayList<>(headers.keySet());
        Map<String, Font> fontMap = new HashedMap();
        Map<Integer, String> datePatternIndexMap = new HashMap<>();
        Map<Integer, Class> dateTypeIndexMap = new HashMap<>();

        //todo 重构Excel导出
        List<MergedRegionSpecification> headerMergedReginSpcifications = new ArrayList<>();
        List<MergedRegionSpecification> dataMergedReginSpcifications = new ArrayList<>();
        List<MergedRegionSpecification> mergedRegionSpecifications = exportParams.getMergedRegionSpecifications();
        if (CollectionUtils.isNotEmpty(mergedRegionSpecifications)) {
            headerMergedReginSpcifications = mergedRegionSpecifications.stream()
                    .filter(MergedRegionSpecification::getIsHeader).collect(Collectors.toList());

            dataMergedReginSpcifications = mergedRegionSpecifications.stream()
                    .filter(m -> !m.getIsHeader()).collect(Collectors.toList());

        }
        if (CollectionUtils.isNotEmpty(headerMergedReginSpcifications)) {
            Ihr360CellStyle defaultHeaderCellStyle = ExcelCellStyleFactory.createDefaultHeaderCellStyle();
            Map<String, CellStyle> excelCellStyleMap = new HashMap<>();
            Map<String, ExcelCellStyle> headerStyleMap = exportParams.getHeaderStyleMap();


            Set<Integer> headerRows = new HashSet<>();
            for (MergedRegionSpecification specification : headerMergedReginSpcifications) {
                ExportHeaderParams exportHeaderParams = specification.getExportHeaderParams();
                if (exportHeaderParams == null) {
                    throw new ExcelException("ExportHeaderParams can not be null!");
                }
                Map<String, String> mergedHeaderMap = exportHeaderParams.getHeaderMap();
                if (MapUtils.isEmpty(mergedHeaderMap)) {
                    throw new ExcelException("mergedHeaderMap can not be empty!");
                }

                Row row = sheet.createRow(specification.getRowNum());
                row.setHeightInPoints(ExcelDefaultConfig.DEFAULT_ROW_HEADER_HEIGHT_INPOINT);
                headerRows.add(specification.getRowNum());

                exportParams.setMergedHeaderMap(mergedHeaderMap);
                datePatternIndexMap = Ihr360ExcelCommonExcelHelper.handleExportHeader(workbook,
                        row, fontMap, exportParams, exportHeaderParams.getStartIndex(), sheet, dateTypeIndexMap);

                //合并单元格
                List<int[]> specifiCationParams = specification.getSpecifiCationParams();

                for (int[] params : specifiCationParams) {
                    sheet.addMergedRegion(new CellRangeAddress(params[0], params[1], params[2], params[3]));
                }
            }
            //合并单元格后再次处理单元格样式
            if (CollectionUtils.isNotEmpty(headerRows)) {
                for (Integer rowNum : headerRows) {
                    Row row = sheet.getRow(rowNum);
                    for (String headerKey : headerKeys) {
                        CellStyle poiCellStyle = getPoiCellStyle(workbook, fontMap, defaultHeaderCellStyle, excelCellStyleMap, headerStyleMap, headerKey);
                        Cell cell = row.getCell(headerKey.indexOf(headerKey));
                        if (cell == null) {
                            continue;
                        }
                        cell.setCellStyle(poiCellStyle);
                    }
                }
            }

        } else {
            // 产生表格标题行
            Row row = sheet.createRow(0);
            row.setHeightInPoints(ExcelDefaultConfig.DEFAULT_ROW_HEADER_HEIGHT_INPOINT);
            datePatternIndexMap = Ihr360ExcelCommonExcelHelper.handleExportHeader(workbook, row, fontMap, exportParams, 0, sheet, dateTypeIndexMap);
        }
        CellStyle textCellStyle = ExcelCellStyleFactory.createdDefaultTextCellStyle(workbook);


        if (CollectionUtils.isNotEmpty(dataMergedReginSpcifications)) {
            Set<Integer> dataRows = new HashSet<>();

            dataMergedReginSpcifications.forEach(dataMergedReginSpcification -> {
                Row row = sheet.createRow(dataMergedReginSpcification.getRowNum());
                dataRows.add(row.getRowNum());
                row.setHeightInPoints(ExcelDefaultConfig.DEFAULT_ROW_HEADER_HEIGHT_INPOINT);
                //合并单元格
                List<int[]> specifiCationParams = dataMergedReginSpcification.getSpecifiCationParams();
                for (int[] params : specifiCationParams) {
                    sheet.addMergedRegion(new CellRangeAddress(params[0], params[1], params[2], params[3]));
                }

                MergedExportData mergedExportData = dataMergedReginSpcification.getExportData();
                Map<String, Object> dataMap = mergedExportData.getDataMap();
                for (int i = 0; i < headerKeys.size(); i++) {
                    Cell cell = row.createCell(i);
                    Ihr360ExcelCellHelper.setCellValue(cell, dataMap.get(headerKeys.get(i)), ExcelDefaultConfig.DEFAULT_OUTPUT_DATE_PATTERN, null, textCellStyle, null);
                }
            });


            CellStyle cellStyle = ExcelCellStyleFactory.createdDefaultTextCellStyle(workbook);
            if (CollectionUtils.isNotEmpty(dataRows)) {
                for (Integer rowNum : dataRows) {
                    Row row = sheet.getRow(rowNum);
                    for (String headerKey : headerKeys) {
                        Cell cell = row.getCell(headerKey.indexOf(headerKey));
                        if (cell == null) {
                            continue;
                        }
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
            // 设定自动宽度
            for (int i = 0; i < headers.size(); i++) {
                setSheetAutoAddWidthColumn(sheet, i);
                sheet.setDefaultColumnStyle(i, textCellStyle);
            }
            return;
        }

        // 遍历集合数据，产生数据行
        if (CollectionUtils.isEmpty(datas)) {

            Ihr360ExcelDropListHelper.handleDropDownList(exportParams, workbook, sheet, 1, 1000);
            for (int i = 0; i < headers.size(); i++) {
                setSheetAutoAddWidthColumn(sheet, i);
                sheet.setDefaultColumnStyle(i, textCellStyle);
            }
            return;
        }
        Ihr360ExcelDropListHelper.handleDropDownList(exportParams, workbook, sheet, 1, datas.size() + 10);


        Iterator<T> datasIt = datas.iterator();
        int index = 0;
        while (datasIt.hasNext()) {
            index++;
            Row row = sheet.getRow(index);
            if (row == null) {
                row = sheet.createRow(index);
            }
            T rowData = datasIt.next();
            try {
                if (rowData instanceof List) {
                    @SuppressWarnings("unchecked")
                    List<Object> cellDatas = (List<Object>) rowData;
                    int cellNum = 0;
                    //遍历列名
                    for (int i = 0; i < cellDatas.size(); i++) {
                        Cell cell = row.createCell(i);
                        String datePattern = datePatternIndexMap.get(i);
                        if (StringUtils.isBlank(datePattern)) {
                            datePattern = ExcelDefaultConfig.DEFAULT_OUTPUT_DATE_PATTERN;
                        }
                        Class clazz = dateTypeIndexMap.get(i);
                        if (clazz != null) {
                            if (clazz.isAssignableFrom(Number.class) || clazz == Date.class) {
                                cell.setCellType(CellType.NUMERIC);
                            }
                        }
                        Ihr360ExcelCellHelper.setCellValue(cell, cellDatas.get(i), datePattern, null, textCellStyle, clazz);
                    }
                } else if (rowData instanceof Map) {
                    Map<String, Object> cellMap = (Map<String, Object>) rowData;
                    for (int i = 0; i < headerKeys.size(); i++) {
                        Cell cell = row.createCell(i);
                        String datePattern = datePatternIndexMap.get(i);
                        if (StringUtils.isBlank(datePattern)) {
                            datePattern = ExcelDefaultConfig.DEFAULT_OUTPUT_DATE_PATTERN;
                        }
                        Ihr360ExcelCellHelper.setCellValue(cell, cellMap.get(headerKeys.get(i)), datePattern, null, textCellStyle, null);
                    }

                } else {
                    List<Field> fields = Ihr360ExcelJavaBeanDataHelper.getExcelCellField(rowData.getClass());
                    Map<String, Field> fieldsMap = new HashedMap();
                    fields.forEach(field -> {
                        ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
                        if (excelCell != null) {
                            String headerKey = excelCell.headerKey();
                            if (StringUtils.isNotBlank(headerKey)) {
                                fieldsMap.put(headerKey, field);
                            }
                        }

                    });
                    for (int i = 0; i < headerKeys.size(); i++) {
                        Cell cell = row.createCell(i);
                        Field field = fieldsMap.get(headerKeys.get(i));
                        field.setAccessible(true);
                        Object value = field.get(rowData);
                        String datePattern = datePatternIndexMap.get(i);
                        if (StringUtils.isBlank(datePattern)) {
                            datePattern = ExcelDefaultConfig.DEFAULT_OUTPUT_DATE_PATTERN;
                        }
                        Ihr360ExcelCellHelper.setCellValue(cell, value, datePattern, field, textCellStyle, null);
                    }
                }
            } catch (Exception e) {
                logger.error(e.toString(), e);
            }
        }
        // 设定自动宽度
        for (int i = 0; i < headers.size(); i++) {
            setSheetAutoAddWidthColumn(sheet, i);
            sheet.setDefaultColumnStyle(i, textCellStyle);
        }


    }


}
