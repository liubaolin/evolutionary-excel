package com.ihr360.excel.util.helper;

import com.ihr360.excel.core.cellstyle.ExcelCellStyle;
import com.ihr360.excel.core.cellstyle.ExcelFont;
import com.ihr360.excel.core.cellstyle.Ihr360CellStyle;
import com.ihr360.excel.core.cellstyle.Ihr360SSCellStyle;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * @author richey
 */
public class Ihr360ExcelCellStyleHelper {

    public static CellStyle getPoiCellStyle(Workbook workbook, Map<String, Font> fontMap, Ihr360CellStyle defaultHeaderCellStyle, Map<String, CellStyle> excelCellStyleMap, Map<String, ExcelCellStyle> headerStyleMap, String headerKey) {
        CellStyle poiCellStyle;
        ExcelCellStyle cellStyle = null;
        if (headerStyleMap != null) {
            cellStyle = headerStyleMap.get(headerKey);
        }
        if (cellStyle != null) {
            poiCellStyle = setExcelCellStyle(cellStyle, workbook, fontMap, excelCellStyleMap);
        } else {
            poiCellStyle = setExcelCellStyle(defaultHeaderCellStyle, workbook, fontMap, excelCellStyleMap);
        }
        return poiCellStyle;
    }

    public static void setSheetAutoAddWidthColumn(Sheet sheet, int i) {
        sheet.autoSizeColumn(i);
        //一个字符的1/256的宽度作为一个单位 最多支持65280单位(255个字符)
        int width = sheet.getColumnWidth(i);
        int addWidth = 256 * 10;
        if ((width + addWidth) > 65280) {
            return;
        }
        sheet.setColumnWidth(i, (width + addWidth));
    }

    public static CellStyle setExcelCellStyle(ExcelCellStyle cellStyle, Workbook workbook, Map<String, Font> fontMap, Map<String, CellStyle> excelCellStyleMap) {

        CellStyle poiCellStyle = null;
        if (cellStyle instanceof Ihr360SSCellStyle) {
            Ihr360SSCellStyle ihr360SSCellStyle = (Ihr360SSCellStyle) cellStyle;
            poiCellStyle = getCellStyle(workbook, excelCellStyleMap, ihr360SSCellStyle);
            setPoiCellStyle(cellStyle, workbook, fontMap, poiCellStyle);
        } else if (cellStyle instanceof Ihr360CellStyle) {
            Ihr360CellStyle ihr360CellStyle = (Ihr360CellStyle) cellStyle;
            poiCellStyle = getCellStyle(workbook, excelCellStyleMap, ihr360CellStyle);
            setPoiCellStyle(cellStyle, workbook, fontMap, poiCellStyle);
        }
        return poiCellStyle;
    }

    public static CellStyle getCellStyle(Workbook workbook, Map<String, CellStyle> excelCellStyleMap, ExcelCellStyle ihr360SSCellStyle) {
        CellStyle poiCellStyle;
        poiCellStyle = excelCellStyleMap.get(ihr360SSCellStyle.getUuid());
        if (poiCellStyle == null) {
            poiCellStyle = workbook.createCellStyle();
            excelCellStyleMap.put(ihr360SSCellStyle.getUuid(), poiCellStyle);
        }
        return poiCellStyle;
    }

    public static void setPoiCellStyle(ExcelCellStyle cellStyle, Workbook workbook, Map<String, Font> fontMap, CellStyle poiCellStyle) {
        ExcelFont excelFont = null;
        //todo ss版本支持

        if (cellStyle instanceof Ihr360CellStyle) {
            Ihr360CellStyle ihr360CellStyle = (Ihr360CellStyle) cellStyle;
            if (ihr360CellStyle.getFillPattern() > 0) {
                poiCellStyle.setFillPattern(ihr360CellStyle.getFillPattern());
            }
            if (ihr360CellStyle.getForegroundColor() > 0) {
                poiCellStyle.setFillForegroundColor(ihr360CellStyle.getForegroundColor());
            }
            if (ihr360CellStyle.getBackgroundColor() > 0) {
                poiCellStyle.setFillBackgroundColor(ihr360CellStyle.getBackgroundColor());
            }

            if (ihr360CellStyle.getBorderBottom() > 0) {
                poiCellStyle.setBorderBottom(ihr360CellStyle.getBorderBottom());
            }
            if (ihr360CellStyle.getBorderLeft() > 0) {
                poiCellStyle.setBorderLeft(ihr360CellStyle.getBorderLeft());
            }
            if (ihr360CellStyle.getBorderTop() > 0) {
                poiCellStyle.setBorderTop(ihr360CellStyle.getBorderTop());
            }
            if (ihr360CellStyle.getBorderRight() > 0) {
                poiCellStyle.setBorderRight(ihr360CellStyle.getBorderRight());
            }
            if (ihr360CellStyle.getHorizontalAlignment() > 0) {
                poiCellStyle.setAlignment(ihr360CellStyle.getHorizontalAlignment());
            }
            if (ihr360CellStyle.getVerticalAlignment() > 0) {
                poiCellStyle.setVerticalAlignment(ihr360CellStyle.getVerticalAlignment());
            }
            excelFont = ihr360CellStyle.getExcelFont();
        }


        if (excelFont != null) {
            String fontUuid = excelFont.getUuid();
            Font poiFont = fontMap.get(fontUuid);
            if (poiFont == null) {
                poiFont = workbook.createFont();
                if (StringUtils.isNotBlank(excelFont.getFontName())) {
                    poiFont.setFontName(excelFont.getFontName());
                }
                if (excelFont.getFontHeightInPoints() > 0) {
                    poiFont.setFontHeightInPoints(excelFont.getFontHeightInPoints());
                }
                poiFont.setItalic(excelFont.getItalic());
                if (excelFont.getColor() > 0) {
                    poiFont.setColor(excelFont.getColor());
                }
                if (excelFont.getUnderline() > 0) {
                    poiFont.setUnderline(excelFont.getUnderline());
                }
                poiFont.setBold(excelFont.getBold());
                poiFont.setStrikeout(excelFont.getStrikeout());
                fontMap.put(fontUuid, poiFont);
            }
            poiCellStyle.setFont(poiFont);
        }
    }

}
