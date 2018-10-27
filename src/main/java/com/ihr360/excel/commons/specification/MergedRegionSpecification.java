package com.ihr360.excel.commons.specification;

import com.ihr360.excel.core.metaData.ExportHeaderParams;
import com.ihr360.excel.core.metaData.MergedExportData;

import java.io.Serializable;
import java.util.List;

/**
 * 单元格合并规格
 */
public class MergedRegionSpecification implements Serializable {

    private static final long serialVersionUID = 6189039192707406210L;

    private int rowNum;

    /**
     * CellRangeAddress的参数集合
     */
    private List<int[]> specifiCationParams;

    private boolean isHeader;

    /**
     * 　对应合并字段的表头
     */
    private ExportHeaderParams exportHeaderParams;

    private MergedExportData exportData;



    private MergedRegionSpecification() {

    }

    private MergedRegionSpecification(int rowNum,List<int[]> specifiCationParams, boolean isHeader) {
        this.rowNum = rowNum;
        this.specifiCationParams = specifiCationParams;
        this.isHeader = isHeader;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public ExportHeaderParams getExportHeaderParams() {
        return exportHeaderParams;
    }

    public void setExportHeaderParams(ExportHeaderParams exportHeaderParams) {
        this.exportHeaderParams = exportHeaderParams;
    }

    public boolean getIsHeader() {
        return isHeader;
    }

    public void setIsHeader(boolean header) {
        isHeader = header;
    }

    public List<int[]> getSpecifiCationParams() {
        return specifiCationParams;
    }

    public void setSpecifiCationParams(List<int[]> specifiCationParams) {
        this.specifiCationParams = specifiCationParams;
    }


    public MergedExportData getExportData() {
        return exportData;
    }

    public void setExportData(MergedExportData exportData) {
        this.exportData = exportData;
    }

    /**
     * 所有的下标都是从0开始
     *
     * @return
     */
    public static MergedRegionSpecification createdSpecification(int rowNum,List<int[]> specifiCationParams,boolean isHeader) {
        return new MergedRegionSpecification(rowNum,specifiCationParams,isHeader);
    }


}
