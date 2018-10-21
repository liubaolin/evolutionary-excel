/**
 * @author SargerasWang
 */
package com.ihr360.excel;

import com.ihr360.excel.metaData.ExportHeaderParams;
import com.ihr360.excel.metaData.ExportParams;
import com.ihr360.excel.specification.MergedRegionSpecification;
import com.ihr360.excel.util.ExcelDateFormatUtil;
import com.ihr360.excel.util.Ihr360ExcelImportUtil;
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
 * The <code>TestExportMap</code>
 *
 * @author richey.liu
 */
public class TestExportMap {

    public static final String EXCEL_TEST_NAME = "excel_test_name";
    public static final String EXCEL_TEST_YEAR = "excel_test_year";
    public static final String EXCEL_TEST_MONTH = "excel_test_month";
    public static final String EXCEL_TEST_SALARY = "excel_test_salary";
    public static final String EXCEL_TEST_TAX = "excel_test_tax";
    public static final String EXCEL_TEST_PAYTIME = "excel_test_payTime";


    @Test
    public void exportXls() throws IOException {
        List<List<Object>> datas = new ArrayList<>();

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


        datas.add(data1);
        datas.add(data2);

        Map<String, String> headerMap = new LinkedHashMap<>();
        headerMap.put("name", "姓名");
        headerMap.put("year", "年");
        headerMap.put("month", "月");
        headerMap.put("salary", "薪资");
        headerMap.put("tax", "税额");
        headerMap.put("excel_test_payTime", "支付时间");

        File f = new File("testExportMapData.xls");
        OutputStream out = new FileOutputStream(f);

   /* //下拉列
    Map<String, List<String>> dropDownsMap = new LinkedHashMap<>();
    List<String> monthDropList = new ArrayList<>();
    monthDropList.add("一月");
    monthDropList.add("二月");
    monthDropList.add("三月");
    monthDropList.add("四月");
    monthDropList.add("五月");
    dropDownsMap.put("month", monthDropList);*/

        //日期列输出格式
        Map<String, String> datePattern = new HashMap<>();
        datePattern.put("excel_test_payTime", ExcelDateFormatUtil.PATTERN_ISO_ON_DATE);


        Map<String, Class> dateTypeParam = new HashMap<>();
        dateTypeParam.put("year", Integer.class);
        dateTypeParam.put("salary", Double.class);



        /**
         * 导出map类型的数据
         * 表头map要用有序的map，导出Excel表头为表头map的顺序
         */

        ExportParams<List<Object>> exportParams = new ExportParams();
        exportParams.setHeaderMap(headerMap);
        exportParams.setRowDatas(datas);
//    exportParams.setDropDownsMap(dropDownsMap);
        exportParams.setDatePatternMap(datePattern);

        exportParams.setDataTypeMap(dateTypeParam);

        Ihr360ExcelImportUtil.exportExcel(exportParams, out);
        out.close();
    }

    /**
     * 　支持合并单元格的导出
     *
     * @throws IOException
     */
    @Test
    public void exportMergeReginXls() throws IOException {

        Map<String, String> headerMap = new LinkedHashMap<>();
        headerMap.put("name", "姓名");
        headerMap.put("department", "部门");
        headerMap.put("annual", "年度");
        headerMap.put("result", "考核结果");
        headerMap.put("remark", "备注");

        List<MergedRegionSpecification> mergedRegionSpecifications = new ArrayList<>();
        List<int[]> firstRowparams = new ArrayList<>();
        firstRowparams.add(new int[]{0, 1, 0, 0});
        firstRowparams.add(new int[]{0, 1, 1, 1});
        firstRowparams.add(new int[]{0, 0, 2, 4});
        MergedRegionSpecification firstMergedSpecification =  MergedRegionSpecification.createdSpecification(0, firstRowparams, true);

        Map<String, String> firstHeaderMap = new LinkedHashMap<>();
        firstHeaderMap.put("name", "姓名");
        firstHeaderMap.put("department", "部门");
        firstHeaderMap.put("annual_assessment", "党团员年度考核1");
        firstHeaderMap.put("annual_assessment", "党团员年度考核1");
        firstHeaderMap.put("annual_assessment", "党团员年度考核1");
        ExportHeaderParams firstHeaderParams = new ExportHeaderParams();
        firstHeaderParams.setHeaderMap(firstHeaderMap);
        firstMergedSpecification.setExportHeaderParams(firstHeaderParams);
        mergedRegionSpecifications.add(firstMergedSpecification);

        List<int[]> secondRowparams = new ArrayList<>();
        secondRowparams.add(new int[]{1, 1, 2, 2});
        secondRowparams.add(new int[]{1, 1, 3, 3});
        secondRowparams.add(new int[]{1, 1, 4, 4});
        MergedRegionSpecification secondMergedSpecification =  MergedRegionSpecification.createdSpecification(1, secondRowparams, true);

        Map<String, String> secondHeaderMap = new LinkedHashMap<>();
        secondHeaderMap.put("annual", "年度");
        secondHeaderMap.put("result", "考核结果");
        secondHeaderMap.put("remark", "备注");
        ExportHeaderParams secondHeaderParams = new ExportHeaderParams();
        secondHeaderParams.setHeaderMap(secondHeaderMap);
        secondHeaderParams.setStartIndex(2);
        secondMergedSpecification.setSpecifiCationParams(secondRowparams);
        secondMergedSpecification.setExportHeaderParams(secondHeaderParams);

        mergedRegionSpecifications.add(secondMergedSpecification);


        File f = new File("testExportMergeRegin.xls");
        OutputStream out = new FileOutputStream(f);
        List<List<Object>> datas = new ArrayList<>();

        ExportParams<List<Object>> exportParams = new ExportParams();
        exportParams.setRowDatas(datas);
        exportParams.setHeaderMap(headerMap);
        exportParams.setMergedRegionSpecifications(mergedRegionSpecifications);
        Ihr360ExcelImportUtil.exportExcel(exportParams, out);
        out.close();

    }

    @Test
    public void exportMergeRegin2Xls() throws IOException {

        Map<String, String> headerMap = new LinkedHashMap<>();


        List<MergedRegionSpecification> mergedRegionSpecifications = new ArrayList<>();
        List<int[]> firstRowparams = new ArrayList<>();
        firstRowparams.add(new int[]{0, 1, 0, 0});
        firstRowparams.add(new int[]{0, 0, 1, 1});
        MergedRegionSpecification firstMergedSpecification =  MergedRegionSpecification.createdSpecification(0, firstRowparams, true);

        Map<String, String> firstHeaderMap = new LinkedHashMap<>();
        firstHeaderMap.put("name", "森马22");
        firstHeaderMap.put("month", "上月");

        ExportHeaderParams firstHeaderParams = new ExportHeaderParams();
        firstHeaderParams.setHeaderMap(firstHeaderMap);
        firstMergedSpecification.setExportHeaderParams(firstHeaderParams);
        mergedRegionSpecifications.add(firstMergedSpecification);

        List<int[]> secondRowparams = new ArrayList<>();
        secondRowparams.add(new int[]{1, 1, 1, 1});

        MergedRegionSpecification secondMergedSpecification =  MergedRegionSpecification.createdSpecification(1, secondRowparams, true);

        Map<String, String> secondHeaderMap = new LinkedHashMap<>();
        secondHeaderMap.put("month", "本月");
        ExportHeaderParams secondHeaderParams = new ExportHeaderParams();
        secondHeaderParams.setHeaderMap(secondHeaderMap);
        secondHeaderParams.setStartIndex(1);
        secondMergedSpecification.setSpecifiCationParams(secondRowparams);
        secondMergedSpecification.setExportHeaderParams(secondHeaderParams);

        mergedRegionSpecifications.add(secondMergedSpecification);


        File f = new File("testExportMergeRegin2.xls");
        OutputStream out = new FileOutputStream(f);
        List<List<Object>> datas = new ArrayList<>();

        ExportParams<List<Object>> exportParams = new ExportParams();
        exportParams.setRowDatas(datas);
        exportParams.setHeaderMap(headerMap);
        exportParams.setMergedRegionSpecifications(mergedRegionSpecifications);
        Ihr360ExcelImportUtil.exportExcel(exportParams, out);
        out.close();

    }

}
