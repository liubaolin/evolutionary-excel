package com.ihr360.excel.util.helper;

import com.ihr360.excel.config.ExcelDefaultConfig;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContextHolder;
import com.ihr360.excel.core.metaData.ImportParams;
import com.ihr360.excel.commons.specification.CommonSpecification;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

/**
 * Excel 规格处理类
 * @author richey
 */
public class Ihr360ExcelSpecificationHelper {

    public static <T> void handleCommonSpecification(Row row, Map<String, Object> mapRow) {
        Ihr360ImportExcelContext<T> excelContext = Ihr360ImportExcelContextHolder.getExcelContext();
        ImportParams importParams = excelContext.getImportParams();
        CommonSpecification commonSpecification = importParams.getCommonSpecification();
        if (commonSpecification != null && commonSpecification.isShowRowNum()) {
            String rowNumKey = StringUtils.isEmpty(commonSpecification.getRowNumKey())
                    ? ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM
                    : commonSpecification.getRowNumKey();
            mapRow.put(rowNumKey, row.getRowNum() + 1);
        }
    }

}
