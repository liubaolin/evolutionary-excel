package top.evolutionary.excel.util.helper;

import top.evolutionary.excel.config.ExcelDefaultConfig;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContext;
import top.evolutionary.excel.commons.context.Ihr360ImportExcelContextHolder;
import top.evolutionary.excel.core.metaData.ImportParams;
import top.evolutionary.excel.commons.specification.CommonSpecification;
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
