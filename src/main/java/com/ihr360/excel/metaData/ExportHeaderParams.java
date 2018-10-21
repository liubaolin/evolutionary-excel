package com.ihr360.excel.metaData;

import java.util.Map;

public class ExportHeaderParams {


    /**
     * 表头要用有序的Map
     */
    private Map<String, String> headerMap;


    private int startIndex;

    public int getStartIndex() {
        return startIndex;
    }

    public void setStartIndex(int startIndex) {
        this.startIndex = startIndex;
    }

    public Map<String, String> getHeaderMap() {
        return headerMap;
    }

    public void setHeaderMap(Map<String, String> headerMap) {
        this.headerMap = headerMap;
    }

}
