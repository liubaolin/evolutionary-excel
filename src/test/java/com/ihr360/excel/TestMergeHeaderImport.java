package com.ihr360.excel;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.ihr360.excel.config.Ihr360TemplateExcelConfiguration;
import com.ihr360.excel.context.Ihr360ImportExcelContext;
import com.ihr360.excel.logs.ExcelRowLog;
import com.ihr360.excel.util.Ihr360ExcelImportUtil;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * 测试合并表头导入
 */
public class TestMergeHeaderImport {

    @Test
    public void importBeanXls() throws FileNotFoundException {
        File f = new File("钉钉考勤月报导出模版.xlsx");
        InputStream inputStream = new FileInputStream(f);
        Ihr360ImportExcelContext ihr360ImportExcelContext = new Ihr360ImportExcelContext();

        Map<String, Integer> testConfiguration = Maps.newLinkedHashMap();

        Ihr360TemplateExcelConfiguration templateExcelConfiguration = Ihr360TemplateExcelConfiguration.builder()
                .templateHeaders(testConfiguration)
                .templateHeaderRowNum(Lists.newArrayList(2, 3))
                .templateDataBeginRowNum(4)
                .checkRepeatHeader(false)
                .importType(Map.class)
                .build();


        Collection<Map> excelDatas = Ihr360ExcelImportUtil.importExcel(templateExcelConfiguration, inputStream, ihr360ImportExcelContext);

        List<ExcelRowLog> excelRowLogs = ihr360ImportExcelContext.getLogs().getRowLogList();

        printExcelDatas(excelDatas);
    }


    private void printExcelDatas(Collection<Map> excelDatas) {
        for (Map rowMap : excelDatas) {
            StringBuilder sb = new StringBuilder();
            rowMap.forEach((key, value) -> {
                sb.append(key).append("：").append(value).append(";");
            });
            System.out.println(sb.toString());
        }
    }


}
