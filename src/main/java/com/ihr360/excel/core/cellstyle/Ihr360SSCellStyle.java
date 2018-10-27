package com.ihr360.excel.core.cellstyle;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.UUID;

public class Ihr360SSCellStyle implements ExcelCellStyle {

    private static final long serialVersionUID = 7375646455614938039L;

    private Ihr360SSCellStyle() {

    }

    /**
     * 用于标识唯一EXcelCellStyle,避免重复创建
     */
    protected String uuid;

    /**
     * 前景色
     */
    protected short foregroundColor;

    /**
     * 背景色
     */
    protected short backgroundColor;



    /**
     * 单元格为水平对齐的类型
     */
    private HorizontalAlignment horizontalAlignment;

    private VerticalAlignment verticalAlignment;

    private ExcelFont excelFont;

    /**
     * 单元格的填充信息模式和纯色填充单元。
     */
    private FillPatternType fillPattern;


    /**
     * 下边框
     */
    private BorderStyle borderBottom;

    /**
     * 左边框
     */
    private BorderStyle borderLeft;

    /**
     * 上边框
     */
    private BorderStyle borderTop;

    /**
     * 右边框
     */
    private BorderStyle borderRight;

    public FillPatternType getFillPattern() {
        return fillPattern;
    }

    public void setFillPattern(FillPatternType fillPattern) {
        this.fillPattern = fillPattern;
    }

    public BorderStyle getBorderBottom() {
        return borderBottom;
    }

    public void setBorderBottom(BorderStyle borderBottom) {
        this.borderBottom = borderBottom;
    }

    public BorderStyle getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(BorderStyle borderLeft) {
        this.borderLeft = borderLeft;
    }

    public BorderStyle getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(BorderStyle borderTop) {
        this.borderTop = borderTop;
    }

    public BorderStyle getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(BorderStyle borderRight) {
        this.borderRight = borderRight;
    }

    private void setUuid(String uuid) {
        this.uuid = uuid;
    }

    public void setForegroundColor(short foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public void setBackgroundColor(short backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public void setExcelFont(ExcelFont excelFont) {
        this.excelFont = excelFont;
    }

    @Override
    public String getUuid() {
        return uuid;
    }
    public short getForegroundColor() {
        return foregroundColor;
    }

    public short getBackgroundColor() {
        return backgroundColor;
    }
    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public ExcelFont getExcelFont() {
        return excelFont;
    }

    public static Ihr360SSCellStyle createExcelCellStyle() {
        Ihr360SSCellStyle excelCellStyle = new Ihr360SSCellStyle();
        excelCellStyle.setUuid(UUID.randomUUID().toString());
        return excelCellStyle;
    }
}
