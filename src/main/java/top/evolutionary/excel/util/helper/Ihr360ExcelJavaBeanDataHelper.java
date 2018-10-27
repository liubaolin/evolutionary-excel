package top.evolutionary.excel.util.helper;

import top.evolutionary.excel.commons.CatcheExcelI18nProps;
import top.evolutionary.excel.core.annotation.ExcelCell;
import top.evolutionary.excel.core.annotation.ExcelConfig;
import top.evolutionary.excel.core.annotation.RowNumberField;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContext;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContextHolder;
import top.evolutionary.excel.commons.exception.ExcelCanHandleException;
import top.evolutionary.excel.commons.exception.ExcelDateParseException;
import top.evolutionary.excel.commons.exception.ExcelException;
import top.evolutionary.excel.commons.exception.ExcellExceptionType;
import top.evolutionary.excel.commons.logs.ExcelLogItem;
import top.evolutionary.excel.commons.logs.ExcelLogType;
import top.evolutionary.excel.core.metaData.ExcelI18nStrategyType;
import top.evolutionary.excel.core.metaData.ImportParams;
import top.evolutionary.excel.commons.specification.ColumnSpecification;
import top.evolutionary.excel.util.date.Ihr360ExcelDateParser;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.MessageFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;
import java.util.stream.Collectors;

/**
 * javaBean类型数数据导入处理
 * @author richey
 */
public class Ihr360ExcelJavaBeanDataHelper {

    private static Logger logger = LoggerFactory.getLogger(Ihr360ExcelJavaBeanDataHelper.class);


    public static <T> T handleImportExcelRowToJavabean(Map<String, Integer> fileHeaderIndexMap, List<ExcelLogItem> rowLogs, Row row) {
        Ihr360ImportExcelContext<T> excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams<T> importParams = excelContext.getImportParams();
        Class<T> clazz = importParams.getImportType();
        Map<String, List<String>> importHeader = importParams.getImportHeader();
        ExcelConfig excelConfig = clazz.getAnnotation(ExcelConfig.class);

        T excelEntityVo = getNewInstance(clazz);

        List<Field> excelCellFields = getExcelCellField(clazz);
        //从传入的表头或者根据i18n策略获得的表头在Excel中对应的列的集合
        List<Integer> configHeaderIndex = new ArrayList<>();

        for (Field field : excelCellFields) {
            RowNumberField rowNumberField = field.getAnnotation(RowNumberField.class);
            if (rowNumberField != null) {
                setEntityFieldValue(clazz, excelEntityVo, field, row.getRowNum() + 1 + "");
                continue;
            }

            String importHeaderKey = field.getAnnotation(ExcelCell.class).headerKey();
            if (StringUtils.isBlank(importHeaderKey)) {
                continue;
            }
            //通过国际化策略得到或者传入的表头
            List<String> configHeaders = getI18nHeadersByStrategy(clazz.getAnnotation(ExcelConfig.class), importHeaderKey);
            if (CollectionUtils.isEmpty(configHeaders) && MapUtils.isNotEmpty(importHeader)) {
                configHeaders = importHeader.get(importHeaderKey);
                if (CollectionUtils.isEmpty(configHeaders)) {
                    //输入表头和国际化策略取的表头均无此表头时忽略不导入
                    continue;
                }
            }

            String currentHeaderName = "";
            Integer cellIndex = null;
            for (String header : configHeaders) {
                cellIndex = fileHeaderIndexMap.get(header);
                if (cellIndex != null) {
                    currentHeaderName = header;
                    break;
                }
            }

            if (Ihr360ExcelValidatorHelper.requiredColumnNotExist(rowLogs, field, cellIndex)) {

                continue;
            }

            configHeaderIndex.add(cellIndex);
            Cell cell = row.getCell(cellIndex);

            boolean valid = Ihr360ExcelValidatorHelper.validateCellData(cell, field, currentHeaderName, rowLogs);

            if (!valid) {
                continue;
            }

            Object excelValueObj = null;
            try {
                excelValueObj = getExcelValueObj(field, cell, excelConfig);
            } catch (ParseException e) {
                rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_CON_NOT_CONVERT_TO_DATE, new String[]{currentHeaderName}));
                continue;
            } catch (NumberFormatException e) {
                rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.COLUMN_FIELD_DATA_TYPE_ERR, new String[]{currentHeaderName}));
                continue;
            }
            if (excelValueObj == null) {
                continue;
            }

            setEntityFieldValue(clazz, excelEntityVo, field, excelValueObj);
        }
        handleImportFlexfieldData(fileHeaderIndexMap, row, configHeaderIndex, importParams, rowLogs, excelEntityVo);

        return excelEntityVo;
    }


    public static <T> T getNewInstance(Class<T> clazz) {
        T excelEntityVo = null;
        try {
            excelEntityVo = clazz.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            throw new ExcelException(MessageFormat.format("can not instance class:{0}", clazz.getSimpleName()), e);
        }
        return excelEntityVo;
    }


    public static List<Field> getExcelCellField(Class<?> clazz) {
        Field[] fieldsArr = clazz.getDeclaredFields();

        Class superClazz = clazz.getSuperclass();
        if (superClazz != null) {
            Field[] superFieldsArr = superClazz.getDeclaredFields();
            fieldsArr = ArrayUtils.addAll(fieldsArr, superFieldsArr);
        }
        if (ArrayUtils.isEmpty(fieldsArr)) {
            return new ArrayList<>();
        }
        return Arrays.stream(fieldsArr)
                .filter(f -> (f.getAnnotation(ExcelCell.class) != null || f.getAnnotation(RowNumberField.class) != null))
                .collect(Collectors.toList());
    }

    public static <T> void setEntityFieldValue(Class<T> clazz, T excelEntityVo, Field flexField, Object fieldValue) {
        flexField.setAccessible(true);
        try {
            flexField.set(excelEntityVo, fieldValue);
        } catch (IllegalAccessException e) {
            throw new ExcelException(MessageFormat.format("Can not set a value {0} for Object: {1} - field: {2}", fieldValue, clazz.getSimpleName(), flexField.getName()), e);
        }
    }

    public static <T> void handleImportFlexfieldData(Map<String, Integer> fileHeaderIndexMap, Row row, List<Integer> configHeaderIndex, ImportParams<T> importParams, List<ExcelLogItem> rowLogs, T excelEntityVo) {

        Class<T> clazz = importParams.getImportType();
        List<Field> excelCellFields = getExcelCellField(clazz);
        List<ColumnSpecification> columnSpecifications = importParams.getColumnSpecifications();
        //列规则信息
        Map<String, ColumnSpecification> columnSpecificationMap = Ihr360ExcelValidatorHelper.getColumnSpecifications(columnSpecifications, fileHeaderIndexMap.keySet());

        Field flexField = getFlexbleField(excelCellFields);
        if (flexField != null) {
            Map<Integer, String> flexFieldHeaderMap = new LinkedHashMap<>();

            fileHeaderIndexMap.forEach((header, index) -> {
                if (!configHeaderIndex.contains(index)) {
                    String flexHeader = flexFieldHeaderMap.get(index);
                    //存在重复表头
                    if (flexHeader != null) {
                        throw new ExcelCanHandleException(ExcellExceptionType.REPEATED_HEADER, new Object[]{index});
                    }
                    flexFieldHeaderMap.put(index, header);
                }
            });
            //有序Map用来存放弹性字段
            Map<String, Object> flexFieldDataMap = new LinkedHashMap<>();

            if (MapUtils.isNotEmpty(flexFieldHeaderMap)) {

                for (Map.Entry<Integer, String> entry : flexFieldHeaderMap.entrySet()) {
                    Integer index = entry.getKey();
                    String header = entry.getValue();
                    Cell cell = row.getCell(index);

                    ColumnSpecification specification = columnSpecificationMap.get(header);
                    if (specification != null && specification.getCellType() != null) {
                        Class type = specification.getCellType();
                        boolean validType = true;
                        if (cell != null) {
                            validType = Ihr360ExcelValidatorHelper.checkBySpecificationType(type, cell);
                        }
                        if (!validType) {
                            rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", header}, index));
                            continue;
                        }
                        try {
                            Ihr360ExcelCellHelper.addCellDataToMap(flexFieldDataMap, header, cell, type);
                        } catch (ParseException e) {
                            rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", header}, index));
                            continue;
                        } catch (ExcelDateParseException | IllegalArgumentException e) {
                            rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", header}, index));
                            continue;
                        }

                    } else {
                        try {
                            Ihr360ExcelCellHelper.addCellDataToMap(flexFieldDataMap, header, cell, null);
                        } catch (ParseException | IllegalArgumentException e) {
                            rowLogs.add(ExcelLogItem.createExcelItem(ExcelLogType.ROW_COLUMN_FIELD_DATA_TYPE_ERR, new String[]{row.getRowNum() + 1 + "", header}, index));
                        }
                    }


                }
                setEntityFieldValue(clazz, excelEntityVo, flexField, flexFieldDataMap);
            }

        }


    }


    public static Field getFlexbleField(List<Field> excelCellFields) {
        Field flexField = null;
        List<Field> flexbleFields = excelCellFields.stream()
                .filter(f -> isFlexbleField(f))
                .collect(Collectors.toList());
        if (flexbleFields.size() > 1) {
            throw new ExcelException("Each object can only have one flexible field，the current object has " + flexbleFields.size());
        } else if (flexbleFields.size() == 1) {
            flexField = flexbleFields.get(0);
        }
        return flexField;
    }

    public static List<String> getI18nHeadersByStrategy(ExcelConfig excelConfig, String importHeaderKey) {
        List<String> i18nHeaders = new ArrayList<>();
        if (excelConfig != null) {
            ExcelI18nStrategyType excelI18nStrategyType = excelConfig.i18nStrategy();
            switch (excelI18nStrategyType) {
                case EXCEL_I18N_STRATEGY_PROPS:
                    String propsFileName = excelConfig.propsFileName();
                    if (StringUtils.isNotBlank(propsFileName)) {
                        try {
                            i18nHeaders = CatcheExcelI18nProps.getI18SortedHeaders(importHeaderKey, propsFileName);
                        } catch (ExecutionException e) {
                            logger.error("根据词条key从配置文件中获取表头时失败，headerKey：" + importHeaderKey + ",propsFileName:" + propsFileName);
                            throw new ExcelException("Error reading Excel configuration file", e.getCause());
                        }
                    }
                    break;
                default:
            }
        }
        return i18nHeaders;
    }

    public static boolean isFlexbleField(Field f) {
        return f.getAnnotation(ExcelCell.class) != null && f.getAnnotation(ExcelCell.class).flexibleField();
    }


    /**
     * 根据字段类型及excelConfig获取单元格value
     *
     * @param field
     * @param cell
     * @param excelConfig
     * @return
     * @throws ParseException
     */
    public static Object getExcelValueObj(Field field, Cell cell, ExcelConfig excelConfig) throws ParseException {
        Object excelValueObj = null;
        if (cell == null || StringUtils.isBlank(String.valueOf(cell))) {
            return excelValueObj;
        }

        Object cellValue = Ihr360ExcelCellHelper.getCellValue(cell, null);
        // String类型的日期转换
        if (Date.class == field.getType() && cell.getCellType() == CellType.STRING) {
            excelValueObj = Ihr360ExcelDateParser.getDate(String.valueOf(cellValue));
        } else if (Timestamp.class == field.getType() && cell.getCellType() == CellType.STRING) {
            excelValueObj = Ihr360ExcelDateParser.getTimestamp(String.valueOf(cellValue));
        } else if (Integer.class == field.getType()) {
            return Integer.parseInt(String.valueOf(cellValue));
        } else if (Double.class == field.getType()) {
            return Double.parseDouble(String.valueOf(cellValue));
        } else if (Long.class == field.getType()) {
            return Long.parseLong(String.valueOf(cellValue));
        } else if (String.class == field.getType()) {
            return String.valueOf(cellValue);
        } else {
            excelValueObj = Ihr360ExcelCellHelper.getCellValue(cell, null);
            // 处理特殊情况,excel的value为String,且bean中为其他,且defaultValue不为空,那就=defaultValue
            ExcelCell annoCell = field.getAnnotation(ExcelCell.class);
            if (excelValueObj instanceof String && !(field.getType() == String.class) && StringUtils.isNotBlank(annoCell.defaultValue())) {
                excelValueObj = annoCell.defaultValue();
            }
        }
        return excelValueObj;
    }


}
