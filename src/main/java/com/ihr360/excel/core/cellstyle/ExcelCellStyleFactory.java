package com.ihr360.excel.core.cellstyle;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public final class ExcelCellStyleFactory {

    public static Ihr360CellStyle createDefaultHeaderCellStyle() {
        Ihr360CellStyle excelCellStyle = createdExcelCellStyle();
        excelCellStyle.setExcelFont(createDefaultHeaderFont());
        return excelCellStyle;
    }


    public static Ihr360CellStyle createRequiredHeaderCellStyle() {
        Ihr360CellStyle excelCellStyle = createdExcelCellStyle();
        ExcelFont excelFont = createDefaultHeaderFont();
        excelFont.setColor(Font.COLOR_RED);
        excelCellStyle.setExcelFont(excelFont);
        return excelCellStyle;
    }

    /**
     * 创建文本格式的CellStyle
     * @param workbook
     * @return
     */
    public static CellStyle createdDefaultTextCellStyle(Workbook workbook){
        CellStyle textCellStyle = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        textCellStyle.setDataFormat(format.getFormat("@"));
        return textCellStyle;
    }


    public static ExcelFont createDefaultHeaderFont() {
        ExcelFont excelFont = ExcelFont.createExcelFont();
        excelFont.setItalic(false);
        excelFont.setColor(Font.COLOR_NORMAL);
        excelFont.setFontHeightInPoints((short) 10);
        excelFont.setBold(true);
        return excelFont;
    }

    /*private static Ihr360SSCellStyle createDefaultExcelCellStyle() {
        Ihr360SSCellStyle excelCellStyle = Ihr360SSCellStyle.createExcelCellStyle();
        //单元格内容水平居中
        excelCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        //浅绿背景色
        excelCellStyle.setForegroundColor(IndexedColors.AQUA.getIndex());
        excelCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //边框
        excelCellStyle.setBorderBottom(BorderStyle.THIN);
        excelCellStyle.setBorderLeft(BorderStyle.THIN);
        excelCellStyle.setBorderTop(BorderStyle.THIN);
        excelCellStyle.setBorderRight(BorderStyle.THIN);
        return excelCellStyle;
    }*/

    private static Ihr360CellStyle createdExcelCellStyle() {
        Ihr360CellStyle excelCellStyle = Ihr360CellStyle.createExcelCellStyle();
        //单元格内容水平居中
        excelCellStyle.setHorizontalAlignment(CellStyle.ALIGN_CENTER);
        excelCellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        //
        excelCellStyle.setForegroundColor(IndexedColors.AQUA.getIndex());
        excelCellStyle.setFillPattern(Ihr360CellStyle.SOLID_FOREGROUND);
        excelCellStyle.setBorderBottom(Ihr360CellStyle.BORDER_THIN);
        excelCellStyle.setBorderLeft(Ihr360CellStyle.BORDER_THIN);
        excelCellStyle.setBorderTop(Ihr360CellStyle.BORDER_THIN);
        excelCellStyle.setBorderRight(Ihr360CellStyle.BORDER_THIN);
        return excelCellStyle;
    }

}
