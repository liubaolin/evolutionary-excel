package com.ihr360.excel.cellstyle;


import java.util.UUID;

/**
 * 兼容poi-11版本的CellStyle
 */
public class Ihr360CellStyle implements ExcelCellStyle {

    private static final long serialVersionUID = -7057746539774507149L;

    public static final short ALIGN_GENERAL = 0;
    public static final short ALIGN_LEFT = 1;
    public static final short ALIGN_CENTER = 2;
    public static final short ALIGN_RIGHT = 3;
    public static final short ALIGN_FILL = 4;
    public static final short ALIGN_JUSTIFY = 5;
    public static final short ALIGN_CENTER_SELECTION = 6;
    public static final short VERTICAL_TOP = 0;
    public static final short VERTICAL_CENTER = 1;
    public static final short VERTICAL_BOTTOM = 2;
    public static final short VERTICAL_JUSTIFY = 3;
    public static final short BORDER_NONE = 0;
    public static final short BORDER_THIN = 1;
    public static final short BORDER_MEDIUM = 2;
    public static final short BORDER_DASHED = 3;
    public static final short BORDER_HAIR = 7;
    public static final short BORDER_THICK = 5;
    public static final short BORDER_DOUBLE = 6;
    public static final short BORDER_DOTTED = 4;
    public static final short BORDER_MEDIUM_DASHED = 8;
    public static final short BORDER_DASH_DOT = 9;
    public static final short BORDER_MEDIUM_DASH_DOT = 10;
    public static final short BORDER_DASH_DOT_DOT = 11;
    public static final short BORDER_MEDIUM_DASH_DOT_DOT = 12;
    public static final short BORDER_SLANTED_DASH_DOT = 13;
    public static final short NO_FILL = 0;
    public static final short SOLID_FOREGROUND = 1;
    public static final short FINE_DOTS = 2;
    public static final short ALT_BARS = 3;
    public static final short SPARSE_DOTS = 4;
    public static final short THICK_HORZ_BANDS = 5;
    public static final short THICK_VERT_BANDS = 6;
    public static final short THICK_BACKWARD_DIAG = 7;
    public static final short THICK_FORWARD_DIAG = 8;
    public static final short BIG_SPOTS = 9;
    public static final short BRICKS = 10;
    public static final short THIN_HORZ_BANDS = 11;
    public static final short THIN_VERT_BANDS = 12;
    public static final short THIN_BACKWARD_DIAG = 13;
    public static final short THIN_FORWARD_DIAG = 14;
    public static final short SQUARES = 15;
    public static final short DIAMONDS = 16;
    public static final short LESS_DOTS = 17;
    public static final short LEAST_DOTS = 18;


    private Ihr360CellStyle() {

    }

    /**
     * 用于标识唯一EXcelCellStyle,避免重复创建
     */
    private String uuid;

    /**
     * 前景色
     */
    private short foregroundColor;

    /**
     * 背景色
     */
    private short backgroundColor;

    /**
     * 单元格为水平对齐的类型
     */
    private short horizontalAlignment;

    /**
     * 垂直对齐类型
     */
    private short verticalAlignment;

    private ExcelFont excelFont;

    /**
     * 单元格的填充信息模式和纯色填充单元。
     */
    private short fillPattern;


    /**
     * 下边框
     */
    private short borderBottom;

    /**
     * 左边框
     */
    private short borderLeft;

    /**
     * 上边框
     */
    private short borderTop;

    /**
     * 右边框
     */
    private short borderRight;

    public short getFillPattern() {
        return fillPattern;
    }

    public void setFillPattern(short fillPattern) {
        this.fillPattern = fillPattern;
    }

    public short getBorderBottom() {
        return borderBottom;
    }

    public void setBorderBottom(short borderBottom) {
        this.borderBottom = borderBottom;
    }

    public short getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(short borderLeft) {
        this.borderLeft = borderLeft;
    }

    public short getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(short borderTop) {
        this.borderTop = borderTop;
    }

    public short getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(short borderRight) {
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


    public void setHorizontalAlignment(short horizontalAlignment) {
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

    public short getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public ExcelFont getExcelFont() {
        return excelFont;
    }

    public short getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(short verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public static Ihr360CellStyle createExcelCellStyle() {
        Ihr360CellStyle excelCellStyle = new Ihr360CellStyle();
        excelCellStyle.setUuid(UUID.randomUUID().toString());
        return excelCellStyle;
    }
}
