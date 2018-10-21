package com.ihr360.excel.metaData;

import com.ihr360.excel.exception.ExcelException;
import org.apache.commons.lang3.ArrayUtils;

import javax.annotation.Nonnull;
import java.io.Serializable;

/**
 * 批注
 */
public class CellComment implements Serializable {

    private static final long serialVersionUID = 2795052910431190903L;
    /**
     * dx1,dy1,dx2,dy2,col1,row1,col2,row2
     * dx1：起始单元格的x偏移量
     * dy1：起始单元格的y偏移量
     * dx2：终止单元格的x偏移量
     * dy2：终止单元格的y偏移量
     * col1：起始单元格列序号，从0开始计算；
     * row1：起始单元格行序号，从0开始计算；
     * col2：终止单元格列序号，从0开始计算；
     * row2：终止单元格行序号，从0开始计算；
     *
     * @return
     */
    private int[] anchorParams;

    private String author;

    private String contentString;

    private boolean isVisible = true;

    public int[] getAnchorParams() {
        return anchorParams;
    }

    public void setAnchorParams(int[] anchorParams) {
        this.anchorParams = anchorParams;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getContentString() {
        return contentString;
    }

    public void setContentString(String contentString) {
        this.contentString = contentString;
    }

    public boolean isVisible() {
        return isVisible;
    }

    public void setVisible(boolean visible) {
        isVisible = visible;
    }

    public static CellComment createCellComment(int[] anchorParams, String author, @Nonnull String contentString,boolean isVisible) {
        if (ArrayUtils.isEmpty(anchorParams) || anchorParams.length != 8) {
            throw new ExcelException("Parameter type error of anchorParams");
        }
        CellComment cellComment = new CellComment();
        cellComment.setAnchorParams(anchorParams);
        cellComment.setAuthor(author);
        cellComment.setContentString(contentString);
        cellComment.setVisible(isVisible);
        return cellComment;
    }

    private static final int[] DEFAULT_ANCHOR_PARAMS = new int[]{255, 125, 1023, 150, 0, 0, 2, 2};

    public static CellComment createDefaultCellComment(@Nonnull String contentString) {
        return createCellComment(DEFAULT_ANCHOR_PARAMS, null, contentString, false);
    }

}
