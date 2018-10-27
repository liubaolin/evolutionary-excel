package com.ihr360.excel.util.helper;

import com.ihr360.excel.core.annotation.ExcelCell;
import com.ihr360.excel.core.annotation.ExcelConfig;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.util.date.Ihr360ExcelDateFormatUtil;
import com.ihr360.excel.util.date.Ihr360ExcelDateParser;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.NumberToTextConverter;

import java.lang.reflect.Field;
import java.sql.Timestamp;
import java.text.ParseException;
import java.util.Date;
import java.util.Map;

/**
 * 单元格处理类
 *
 * @author richey
 */
public class Ihr360ExcelCellHelper {

    /**
     * Cell是null 或是 CellType.STRING类型的，但是值是blank的
     *
     * @param cell
     * @return
     */
    public static boolean isNullOrBlankStringCell(Cell cell) {
        return cell == null
                || cell.getCellType() == Cell.CELL_TYPE_BLANK
                || (cell.getCellType() == Cell.CELL_TYPE_STRING && StringUtils.isBlank(cell.getStringCellValue()));
    }

    /**
     * 获取单元格值
     *
     * @param cell
     * @param cellValueType
     * @return
     */
    public static Object getCellValue(Cell cell, Class cellValueType) throws ParseException {

        if (isNullOrBlankStringCell(cell)) {
            return null;
        }
        Ihr360ImportExcelContext excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ExcelConfig excelConfig = excelContext.getExcelConfig();

        int cellType = cell.getCellType();
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK:
                return null;
            case Cell.CELL_TYPE_BOOLEAN:
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_ERROR:
                return cell.getErrorCellValue();
            case Cell.CELL_TYPE_FORMULA:
                //公式四舍五入保留两位
                try {
                    return new java.text.DecimalFormat("#.00").format(cell.getNumericCellValue());
                } catch (Exception e) {
                    return "";
                }
            case Cell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {// 判断是日期类型
                    if (cellValueType == String.class) {
                        DataFormatter formatter = Ihr360ExcelCommonExcelHelper.getDataFormatter();
                        return formatter.formatCellValue(cell);
                    }
                    return cell.getDateCellValue();
                } else {
                    // 解决问题：
                    // 1、科学计数法(如2.6E+10)，
                    // 2、超长小数小数位不一致（如1091.19649281798读取出1091.1964928179796），
                    // 3、整型变小数（如0读取出0.0）

                    String strValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                    return getCellValue(cellValueType, strValue);

                }

            case Cell.CELL_TYPE_STRING:
                String cellValue = cell.getStringCellValue();
                if (StringUtils.isNotBlank(cellValue) && excelConfig != null && StringUtils.isNotEmpty(excelConfig.globalRemovePattern())) {
                    cellValue = cellValue.trim().replaceAll(excelConfig.globalRemovePattern(), "");
                } else if (StringUtils.isEmpty(cellValue)) {
                    return null;
                }
                return getCellValue(cellValueType, cellValue);
            default:
                return null;
        }
    }

    private static Object getCellValue(Class specificationType, String cellValue) throws ParseException {

        if (specificationType == null) {
            return cellValue;
        }
        if (specificationType == Date.class) {
            return Ihr360ExcelDateParser.getDate(cellValue);
        } else if (specificationType == Timestamp.class) {
            return Ihr360ExcelDateParser.getTimestamp(cellValue);
        } else if (specificationType == Double.class) {
            return Double.parseDouble(cellValue);
        } else if (specificationType == Long.class) {
            return Long.parseLong(cellValue);
        } else if (specificationType == Integer.class) {
            return Integer.parseInt(cellValue);
        } else {
            return cellValue;
        }
    }

    public static void setCellValue(Cell cell, Object value, String pattern, Field field, CellStyle textCellStyle, Class clazz) {
        String textValue = null;


        if (value instanceof Integer || Integer.class == clazz) {
            String tempStrValue = value == null ? StringUtils.EMPTY : value.toString();
            if (StringUtils.isNotBlank(tempStrValue)) {
                Integer integerValue = Integer.parseInt(tempStrValue);
                cell.setCellValue(integerValue.intValue());
            } else {
                cell.setCellValue(StringUtils.EMPTY);
            }

        } else if (value instanceof Float || Float.class == clazz) {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            String tempStrValue = value == null ? StringUtils.EMPTY : value.toString();
            if (StringUtils.isNotBlank(tempStrValue)) {
                Float floatValue = Float.parseFloat(tempStrValue);
                cell.setCellValue(floatValue.floatValue());
            } else {
                cell.setCellValue(StringUtils.EMPTY);
            }
        } else if (value instanceof Double || Double.class == clazz) {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            String tempStrValue = value == null ? StringUtils.EMPTY : value.toString();
            if (StringUtils.isNotBlank(tempStrValue)) {
                Double doubleValue = Double.parseDouble(tempStrValue);
                cell.setCellValue(doubleValue.doubleValue());
            } else {
                cell.setCellValue(StringUtils.EMPTY);
            }

        } else if (value instanceof Long || Long.class == clazz) {

            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            String tempStrValue = value == null ? StringUtils.EMPTY : value.toString();
            if (StringUtils.isNotBlank(tempStrValue)) {
                Long langValue = Long.parseLong(tempStrValue);
                cell.setCellValue(langValue.longValue());
            } else {
                cell.setCellValue(StringUtils.EMPTY);
            }

        } else if (value instanceof Boolean || Boolean.class == clazz) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            textValue = Ihr360ExcelDateFormatUtil.formatDate(pattern, date.getTime());
        }
        //TODO 暂不支持数组导出
        /*else if (value instanceof String[]) {
            String[] strArr = (String[]) value;
            for (int j = 0; j < strArr.length; j++) {
                String str = strArr[j];
                cell.setCellValue(str);
                if (j != strArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        } else if (value instanceof Double[]) {
            Double[] douArr = (Double[]) value;
            for (int j = 0; j < douArr.length; j++) {
                Double val = douArr[j];
                // 值不为空则set Value
                if (val != null) {
                    cell.setCellValue(val);
                }

                if (j != douArr.length - 1) {
                    cellNum++;
                    cell = row.createCell(cellNum);
                }
            }
        }*/
        else {
            // 其它数据类型都当作字符串简单处理
            String defaultStr = StringUtils.EMPTY;
            if (field != null) {
                ExcelCell anno = field.getAnnotation(ExcelCell.class);
                if (anno != null) {
                    defaultStr = anno.defaultValue();
                }
            }
            textValue = value == null ? defaultStr : value.toString();
        }
        if (textValue != null) {
            if (NumberUtils.isCreatable(textValue)) {
                NumberUtils.createNumber(textValue);
                cell.setCellType(Cell.CELL_TYPE_NUMERIC);
                cell.setCellValue(NumberUtils.createDouble(textValue));
            } else {
                HSSFRichTextString richString = new HSSFRichTextString(textValue);
                cell.setCellStyle(textCellStyle);
                cell.setCellValue(richString);
            }
        }
    }


    public static void addCellDataToMap(Map<String, Object> flexFieldDataMap, String header, Cell cell, Class specificationType) throws ParseException {
        if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            flexFieldDataMap.put(header, null);
        } else {
            Object value = Ihr360ExcelCellHelper.getCellValue(cell, specificationType);

            flexFieldDataMap.put(header, value);

        }
    }


}
