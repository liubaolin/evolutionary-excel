package com.ihr360.excel.handler;

import com.ihr360.excel.annotation.ExcelConfig;
import com.ihr360.excel.context.Ihr360ImportExcelContext;
import com.ihr360.excel.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.exception.ExcelCanHandleException;
import com.ihr360.excel.exception.ExcelException;
import com.ihr360.excel.exception.ExcellExceptionType;
import com.ihr360.excel.logs.ExcelLogItem;
import com.ihr360.excel.logs.ExcelLogType;
import com.ihr360.excel.metaData.ImportParams;
import com.ihr360.excel.specification.ColumnSpecification;
import com.ihr360.excel.specification.CommonSpecification;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.text.ParseException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author richey
 */
public class Ihr360ExcelRowUtil {

    public static boolean checkBlankRow(Row row) {
        boolean allRowIsNull = true;
        if (row == null) {
            return allRowIsNull;
        }
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Object cellValue;
            try {
                cellValue = Ihr360ExcelCellHandler.getCellValue(cellIterator.next(), null);
            } catch (ParseException e) {
                throw new ExcelException("验证空行时，发生数据转换异常！");
            }
            if (cellValue != null && StringUtils.isNotBlank(String.valueOf(cellValue))) {
                allRowIsNull = false;
                break;
            }
        }
        return allRowIsNull;
    }

    public static Map<String, Object> handleExcelRowToMap(Map<String, Integer> fileHeaderIndexMap, Row row, List<ExcelLogItem> rowLogs, List<ColumnSpecification> columnSpecifications) {
        Map<String, Object> map = new LinkedHashMap<>();
        // 判空
        for (Map.Entry<String, Integer> entry : fileHeaderIndexMap.entrySet()) {
            String fileHeaderName = entry.getKey();
            Integer headerIndex = fileHeaderIndexMap.get(fileHeaderName);
            Cell cell = row.getCell(headerIndex);

            //列规则信息
            Map<String, ColumnSpecification> columnSpecificationMap = Ihr360ExcelValidatorHandler.getColumnSpecifications(columnSpecifications, fileHeaderIndexMap.keySet());
            ColumnSpecification specification = columnSpecificationMap.get(fileHeaderName);
            boolean checkSpecification = specification != null && specification.getCellType() != null;
            if (checkSpecification) {
                if (Ihr360ExcelCellHandler.isNullOrBlankStringCell(cell) && !specification.isAllowNull()) {
                    rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_DATA_REQUIRED, new String[]{fileHeaderName}, headerIndex));
                    continue;
                } else if (Ihr360ExcelCellHandler.isNullOrBlankStringCell(cell) && specification.isAllowNull()) {

                    try {
                        Ihr360ExcelCellHandler.addCellDataToMap(map, fileHeaderName, cell, specification.getCellType());
                    } catch (ParseException e) {
                        rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", fileHeaderName}, headerIndex));
                    }
                    continue;
                }

                Class type = specification.getCellType();
                boolean validType = Ihr360ExcelValidatorHandler.checkBySpecificationType(type, cell);
                if (!validType) {
                    rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", fileHeaderName}, headerIndex));
                    continue;
                }
                try {
                    Ihr360ExcelCellHandler.addCellDataToMap(map, fileHeaderName, cell, type);
                } catch (ParseException | IllegalArgumentException e) {
                    rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", fileHeaderName}, headerIndex));
                    continue;
                }

            } else {
                try {
                    Ihr360ExcelCellHandler.addCellDataToMap(map, fileHeaderName, cell, null);
                } catch (ParseException | IllegalArgumentException e) {
                    rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", fileHeaderName}, headerIndex));
                }
            }
        }

        return map;
    }


    /**
     * row转Map<列名，列index>
     *
     * @param row
     * @return
     */
    public static Map<String, Integer> convertRowToHeaderMap(Row row) {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        CommonSpecification commonSpecification = excelContext.getImportParams().getCommonSpecification();
        boolean checkRepeatHeader = commonSpecification == null ? true : commonSpecification.isCheckRepeatHeader();

        Map<String, Integer> headerTitleIndexMap = new LinkedHashMap<>();
        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String headerTitle;

            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    // 判断是日期类型
                    if (DateUtil.isCellDateFormatted(cell)) {
                        throw new ExcelCanHandleException(ExcellExceptionType.HEADER_DATE_FIELD_NOTSUPPORT);
                    } else {
                        headerTitle = NumberToTextConverter.toText(cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    headerTitle = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BLANK:
                    continue;
                default:
                    throw new ExcelCanHandleException(ExcellExceptionType.HEADER_DATA_TYPE_NOTSUPPORT);
            }

            //存在重复表头
            if (checkRepeatHeader) {
                for (String exitHeaderTitle : headerTitleIndexMap.keySet()) {
                    if (Ihr360ExcelValidatorHandler.headerEqueals(exitHeaderTitle, headerTitle)) {
                        throw new ExcelCanHandleException(ExcellExceptionType.REPEATED_HEADER, new Object[]{headerTitle});
                    }
                }
            }


            headerTitleIndexMap.put(headerTitle, cell.getColumnIndex());
        }
        return headerTitleIndexMap;
    }

    public static boolean ignorImportRow(Row row) {


        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams importParams = excelContext.getImportParams();

        CommonSpecification commonSpecification = importParams.getCommonSpecification();
        if (commonSpecification == null || MapUtils.isEmpty(commonSpecification.getTemplateHeaders())) {
            return false;
        }
        if (commonSpecification.getTemplateDataBeginRowNum() == null) {
            return false;
        }

        if (row.getRowNum() >= commonSpecification.getTemplateDataBeginRowNum()) {
            return false;
        }




        return true;
    }

    public static boolean ignorHiddenRows() {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ExcelConfig excelConfig = excelContext.getExcelConfig();
        return excelConfig == null || excelConfig.ignoreHiddenRows();
    }

    public static boolean isTemplateHeaderRow(Row row) {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        CommonSpecification commonSpecification = excelContext.getImportParams().getCommonSpecification();

        if (commonSpecification == null) {
            return false;
        }

        if (CollectionUtils.isEmpty(commonSpecification.getTemplateHeaderRowNums())) {
            return false;
        }

        if (!commonSpecification.getTemplateHeaderRowNums().contains(row.getRowNum())) {
            return false;
        }

        return true;
    }

    public static boolean isHeaderRow(Row headerRow, Row row) {
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        CommonSpecification commonSpecification = excelContext.getImportParams().getCommonSpecification();

        List<List<String>> headerJudgeList = null;
        if (commonSpecification != null) {
            headerJudgeList = commonSpecification.getHeaderColumnJudge();
        }

        if (row.getRowNum() != 0 && CollectionUtils.isEmpty(headerJudgeList)) {
            return false;
        }
        if (headerRow != null) {
            return false;
        }
        if (commonSpecification != null && CollectionUtils.isNotEmpty(commonSpecification.getTemplateHeaderRowNums())) {
            return false;
        }
        return true;
    }


    public static boolean isHiddenOrBlanRow(Row row) {
        return row.getZeroHeight() || checkBlankRow(row);
    }
}
