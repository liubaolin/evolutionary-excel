package com.ihr360.excel.metaData;

import java.util.Map;

/**
 * @author richey
 */
public class MergedExportData {

    /**
     * 表头要用有序的Map
     */
    private Map<String, Object> dataMap;

    private int startIndex;


    public MergedExportData(Map<String, Object> dataMap, int startIndex) {
        this.dataMap = dataMap;
        this.startIndex = startIndex;
    }

    public Map<String, Object> getDataMap() {
        return dataMap;
    }

    public void setDataMap(Map<String, Object> dataMap) {
        this.dataMap = dataMap;
    }

    public int getStartIndex() {
        return startIndex;
    }

    public void setStartIndex(int startIndex) {
        this.startIndex = startIndex;
    }
}
