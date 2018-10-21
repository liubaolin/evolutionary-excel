package com.ihr360.excel.handler;

import com.ihr360.excel.constants.ExcelDefaultConfig;
import com.ihr360.excel.metaData.ImportParams;
import com.ihr360.excel.specification.CommonSpecification;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

/**
 * Excel 规格处理类
 * @author richey
 */
public class Ihr360ExcelSpecificationHandler {

    public static <T> void handleCommonSpecification(ImportParams<T> importParams, Row row, Map<String, Object> map) {
        CommonSpecification commonSpecification = importParams.getCommonSpecification();
        if (commonSpecification != null && commonSpecification.isShowRowNum()) {
            String rowNumKey = StringUtils.isEmpty(commonSpecification.getRowNumKey())
                    ? ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM
                    : commonSpecification.getRowNumKey();
            map.put(rowNumKey, row.getRowNum() + 1);
        }
    }

}
