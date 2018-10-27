package com.ihr360.excel.util.helper;

import com.ihr360.excel.core.annotation.ExcelCell;
import com.ihr360.excel.config.ExcelDefaultConfig;
import com.ihr360.excel.commons.exception.ExcelException;
import com.ihr360.excel.commons.logs.ExcelLogItem;
import com.ihr360.excel.commons.logs.ExcelLogType;
import com.ihr360.excel.core.metaData.CellTypeMode;
import com.ihr360.excel.commons.specification.ColumnSpecification;
import com.ihr360.excel.util.date.Ihr360ExcelDateParser;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Field;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @author richey
 */
public class Ihr360ExcelValidatorHelper {

    private static final Pattern NUMBER_PATTERN = Pattern.compile("[1-9]+[0-9]*(\\.[0-9]+)*");

    //非空列不存在
    public static boolean requiredColumnNotExist(List<ExcelLogItem> rowLogs, Field field, Integer cellIndex) {
        boolean nextField = false;
        boolean columnNotExist = (cellIndex == null || cellIndex < 0);
        if (columnNotExist) {
            nextField = true;
            ExcelCell.Valid excelValid = field.getAnnotation(ExcelCell.Valid.class);
            if (excelValid != null && !excelValid.allowNull()) {
                ExcelCell excelCell = field.getAnnotation(ExcelCell.class);
                String defaultHeaderName = "";
                if (excelCell != null) {
                    defaultHeaderName = excelCell.defaultHeaderName();
                }
                rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.REQUIRED_COLUMN_HEADER_NOT_FOUND, new String[]{defaultHeaderName}));
            }
        }
        return nextField;
    }

    /**
     * 校验Cell数据是否正确
     *
     * @param cell             cell单元格
     * @param field            字段
     * @param localeHeaderName 当前列的表头
     * @return
     */
    public static boolean validateCellData(Cell cell, Field field, String localeHeaderName, List<ExcelLogItem> logItems) {

        ExcelCell annoCell = field.getAnnotation(ExcelCell.class);
        if (annoCell == null) {
            return true;
        }

        CellTypeMode cellTypeMode = annoCell.cellTypeMode();
        Integer[] cellTypeArr = null;

        if (CellTypeMode.LOOSE == cellTypeMode) {
            cellTypeArr = ExcelDefaultConfig.looseValidateMap.get(field.getType());
        }

        if (cellTypeArr == null) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.UNSUPPORTED_TYPE, new String[]{field.getType().getSimpleName()}));
            return false;
        }

        ExcelCell.Valid excelValid = field.getAnnotation(ExcelCell.Valid.class);
        if (excelValid == null) {
            return true;
        }
        if (!cellNullValid(cell, excelValid.allowNull())) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_DATA_REQUIRED, new String[]{localeHeaderName}));
            return false;

        }
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return true;
        }

        List<Integer> cellTypes = Arrays.asList(cellTypeArr);
        // 如果了类型不在指定范围内
        if (!cellTypes.contains(cell.getCellType())) {
            handleCellTypeLog(localeHeaderName, cellTypes, logItems);
            return false;
        }

        // 类型符合验证,但值不在要求范围内的
        // String in
        if (excelValid.in().length != 0 && cell.getCellType() == Cell.CELL_TYPE_STRING) {
            return checkExcelValidIn(cell, localeHeaderName, excelValid, logItems);
        }
        // 数值型 或 可以转为数值的String
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC || cell.getCellType() == Cell.CELL_TYPE_STRING) {
            double cellValue;
            if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                String cellValueStr = cell.getStringCellValue();
                if (matchNumber(cellValueStr)) {
                    cellValue = Double.parseDouble(cellValueStr);
                } else {
                    return true;
                }

            } else {
                cellValue = cell.getNumericCellValue();
            }
            return checkExcelValidNumberRange(localeHeaderName, excelValid, cellValue, logItems);
        }
        return true;
    }

    public static boolean checkExcelValidNumberRange(String localeHeaderName, ExcelCell.Valid excelValid, double cellValue, List<ExcelLogItem> logItems) {
        boolean result = true;
        // 小于
        if (!Double.isNaN(excelValid.lt()) && !(cellValue < excelValid.lt())) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_SCOPE_LT, new Object[]{localeHeaderName, excelValid.lt()}));
            result = false;
        }
        // 大于
        if (!Double.isNaN(excelValid.gt()) && !(cellValue > excelValid.gt())) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_SCOPE_GT, new Object[]{localeHeaderName, excelValid.gt()}));
            result = false;
        }
        // 小于等于
        if (!Double.isNaN(excelValid.le()) && !(cellValue <= excelValid.le())) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_SCOPE_LE, new Object[]{localeHeaderName, excelValid.le()}));
            result = false;

        }
        // 大于等于
        if (!Double.isNaN(excelValid.ge()) && !(cellValue >= excelValid.ge())) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_SCOPE_GE, new Object[]{localeHeaderName, excelValid.ge()}));
            result = false;
        }
        return result;
    }

    public static boolean checkExcelValidIn(Cell cell, String localeHeaderName, ExcelCell.Valid excelValid, List<ExcelLogItem> logItems) {
        boolean result = true;
        String[] in = excelValid.in();
        String cellValue = cell.getStringCellValue();
        boolean isIn = false;
        for (String str : in) {
            if (str.equals(cellValue)) {
                isIn = true;
            }
        }
        if (!isIn) {
            logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_IN_SCOPE, new Object[]{localeHeaderName, in}));
            result = false;
        }
        return result;
    }

    public static Map<String, ColumnSpecification> getColumnSpecifications(List<ColumnSpecification> columnSpecifications, Set<String> headers) {
        Map<String, ColumnSpecification> columnSpecificationMap = new HashMap<>();
        if (CollectionUtils.isNotEmpty(columnSpecifications)) {
            for (ColumnSpecification columnSpecification : columnSpecifications) {
                List<String> columns = columnSpecification.getColumns();
                if (CollectionUtils.isEmpty(columns) && !columnSpecification.getIgnoreColumn()) {
                    continue;
                } else {
                    if (columnSpecification.getIgnoreColumn() && CollectionUtils.isNotEmpty(headers)) {
                        List<String> ignoreColumns = columnSpecification.getColumns();
                        if (CollectionUtils.isEmpty(ignoreColumns)) {
                            columns = new ArrayList<>(headers);
                        } else {
                            columns = headers.stream().filter(header -> !ignoreColumns.contains(header)).collect(Collectors.toList());
                        }
                    }
                }
                columns.forEach(column -> {
                    ColumnSpecification specification = columnSpecificationMap.get(column);
                    if (specification != null) {
                        throw new ExcelException(MessageFormat.format("请检查ColumnSpecification设置，{0}存在多个ColumnSpecification配置", column));
                    }
                    columnSpecificationMap.put(column, columnSpecification);

                });
            }
        }
        return columnSpecificationMap;
    }


    //校验cell为空或""的合法性
    public static boolean cellNullValid(Cell cell, boolean allowNull) {
        return !Ihr360ExcelCellHelper.isNullOrBlankStringCell(cell) || (Ihr360ExcelCellHelper.isNullOrBlankStringCell(cell) && allowNull);
    }

    public static void handleCellTypeLog(String localeHeaderName, List<Integer> cellTypes, List<ExcelLogItem> logItems) {
        StringBuilder strType = new StringBuilder();
        for (int i = 0; i < cellTypes.size(); i++) {
            Integer cellType = cellTypes.get(i);
            strType.append(getCellTypeByInt(cellType));
            if (i != cellTypes.size() - 1) {
                strType.append(",");
            }
        }
        logItems.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_TYPE_CONSTRAINT, new String[]{localeHeaderName, strType.toString()}));
    }

    public static boolean matchNumber(String cellValueStr) {
        Matcher nummMatcher = NUMBER_PATTERN.matcher(cellValueStr);
        return nummMatcher.matches();
    }

    /**
     * 获取cell类型的文字描述
     *
     * @param cellType <pre>
     *                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 @return
     */
    public static String getCellTypeByInt(int cellType) {
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK:
                return "Null type";
            case Cell.CELL_TYPE_BOOLEAN:
                return "Boolean type";
            case Cell.CELL_TYPE_ERROR:
                return "Error type";
            case Cell.CELL_TYPE_FORMULA:
                return "Formula type";
            case Cell.CELL_TYPE_NUMERIC:
                return "Numeric type or String type or Date type";
            case Cell.CELL_TYPE_STRING:
                return "String type or Numeric type or String type or Date type";
            default:
                return "Unknown type";
        }
    }

    /**
     * 根据specification校验Cell数据类型
     *
     * @param clazz
     * @param cell
     */
    public static boolean checkBySpecificationType(Class clazz, Cell cell) {

        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK:
                return true;
            case Cell.CELL_TYPE_BOOLEAN:
                if (clazz != Boolean.class) {
                    return false;
                }
            case Cell.CELL_TYPE_ERROR:
                return true;
            case Cell.CELL_TYPE_FORMULA:
                return true;
            case Cell.CELL_TYPE_NUMERIC:
                if (clazz != Double.class
                        && clazz != Long.class
                        && clazz != Integer.class
                        && clazz != String.class
                        && clazz != Date.class
                        && clazz != Timestamp.class
                        && clazz != Time.class) {
                    return false;
                }
                return true;
            case Cell.CELL_TYPE_STRING:
                if (clazz == Date.class
                        || clazz == Timestamp.class
                        || clazz == Time.class) {
                    String pattern = Ihr360ExcelDateParser.determineDateFormat(cell.getStringCellValue());
                    if (StringUtils.isBlank(pattern)) {
                        return false;
                    }
                }
                return true;
            default:
                return true;
        }
    }

    public static boolean judgeHeader(Map<String, Integer> fileHeaderIndexMap, List<List<String>> headerJudgeList) {

        if (CollectionUtils.isEmpty(headerJudgeList)) {
            return true;
        }

        boolean containsHeader = false;
        Set<String> headerSet = fileHeaderIndexMap.keySet();
        for (List<String> ailiasHeaders : headerJudgeList) {
            containsHeader = false;
            if (CollectionUtils.isEmpty(ailiasHeaders)) {
                continue;
            }

            //不区分大小写
            for (String ailiasHeader : ailiasHeaders) {
                for (String header : headerSet) {
                    if (headerEqueals(header, ailiasHeader)) {
                        containsHeader = true;
                    }
                }
            }
            if (!containsHeader) {
                break;
            }
        }

        return containsHeader;
    }

    public static boolean headerEqueals(String sourceHeader, String targetHeader) {
        if (sourceHeader == null || targetHeader == null) {
            return false;
        }
        return StringUtils.equalsAnyIgnoreCase(sourceHeader.replaceAll("\\s*", ""), targetHeader.replaceAll("\\s*", ""));
    }

}
