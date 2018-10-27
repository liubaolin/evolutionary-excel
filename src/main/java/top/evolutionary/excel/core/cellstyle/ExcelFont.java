package top.evolutionary.excel.core.cellstyle;

import java.util.UUID;

public class ExcelFont {

    /**
     * 用于标识唯一EXcelFont,避免Wookboook重复创建Font
     */
    private String uuid;

    /**
     * 字体名字
     */
    private String fontName;
    /**
     * 字体大小
     */
    private short fontHeightInPoints;

    /**
     * 是否斜体
     */
    private boolean italic = false;

    /**
     * 字体颜色
     */
    private short color;

    /**
     * 下划线
     */
    private byte underline;

    /**
     * 是否加粗
     */
    private boolean bold;

    /**
     * 是否加删除线
     */
    private boolean strikeout = false;

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    public boolean getItalic() {
        return italic;
    }

    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public byte getUnderline() {
        return underline;
    }

    public void setUnderline(byte underline) {
        this.underline = underline;
    }

    public boolean getBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public boolean getStrikeout() {
        return strikeout;
    }

    public void setStrikeout(boolean strikeout) {
        this.strikeout = strikeout;
    }

    public String getUuid() {
        return uuid;
    }

    private void setUuid(String uuid) {
        this.uuid = uuid;
    }

    private ExcelFont() {
    }

    public static ExcelFont createExcelFont(){
        ExcelFont excelFont = new ExcelFont();
        excelFont.setUuid(UUID.randomUUID().toString());
        return excelFont;
    }
}
