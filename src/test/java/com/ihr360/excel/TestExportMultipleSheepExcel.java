package com.ihr360.excel;

import com.google.common.collect.Lists;
import com.ihr360.excel.core.metaData.ExcelSheet;
import com.ihr360.excel.core.metaData.ExportParams;
import com.ihr360.excel.util.Ihr360ExcelExportUtil;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 多sheet导出测试
 *
 * @author richey.liu
 */
public class  TestExportMultipleSheepExcel {

    @Test
    public void export() throws IOException {
        File f = new File("testExportMultipleSheet.xls");
        OutputStream out = new FileOutputStream(f);

        ExportParams<List<Object>> exportParams = new ExportParams();


        List<ExcelSheet<List<Object>>> sheets = Lists.newArrayList();


        Map<String, String> headerMap = new LinkedHashMap<>();
        headerMap.put("name", "姓名");
        headerMap.put("year", "年");
        headerMap.put("month", "月");
        headerMap.put("salary", "薪资");
        headerMap.put("tax", "税额");
        headerMap.put("excel_test_payTime", "支付时间");

        Map<String, Class> dateTypeParam = new HashMap<>();
        dateTypeParam.put("year", Integer.class);
        dateTypeParam.put("salary", Double.class);



        List<List<Object>> sheet1Datas = new ArrayList<>();

        List<Object> data1 = new ArrayList<>();
        data1.add("Foo");
        data1.add("2017");
        data1.add("12");
        data1.add(null);
        data1.add(123);
        data1.add(new Date());
        List<Object> data2 = new ArrayList<>();
        data2.add("Hoo");
        data2.add(2017);
        data2.add(11);
        data2.add(1563.23);
        data2.add(125.14);
        data2.add(new Date());
        sheet1Datas.add(data1);
        sheet1Datas.add(data2);


        List<List<Object>> sheet2Datas = new ArrayList<>();

        List<Object> sheet2Data1 = new ArrayList<>();
        sheet2Data1.add("Foo");
        sheet2Data1.add("2017");
        sheet2Data1.add("12");
        sheet2Data1.add(null);
        sheet2Data1.add(123);
        sheet2Data1.add(new Date());
        List<Object> sheet2Data2 = new ArrayList<>();
        sheet2Data2.add("Hoo");
        sheet2Data2.add(2017);
        sheet2Data2.add(11);
        sheet2Data2.add(1563.23);
        sheet2Data2.add(125.14);
        sheet2Data2.add(new Date());
        sheet2Datas.add(sheet2Data1);
        sheet2Datas.add(sheet2Data2);

        ExcelSheet<List<Object>> sheet1 = new ExcelSheet<List<Object>>();
        sheets.add(sheet1);
        sheet1.setSheetName("第一个sheet");
        sheet1.setHeaders(headerMap);
        sheet1.setDataset(sheet1Datas);
        sheet1.setDataTypeMap(dateTypeParam);

        ExcelSheet<List<Object>> sheet2 = new ExcelSheet<List<Object>>();
        sheets.add(sheet2);
        sheet2.setSheetName("第二个sheet");
        sheet2.setHeaders(headerMap);
        sheet2.setDataset(sheet2Datas);
//        sheet2.setDataTypeMap(dateTypeParam);

        exportParams.setSheets(sheets);

        Ihr360ExcelExportUtil.exportExcel(exportParams, out);
        out.close();
    }

}
