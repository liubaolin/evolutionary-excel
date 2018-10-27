package com.ihr360.excel;

import com.ihr360.excel.core.cellstyle.ExcelCellStyle;
import com.ihr360.excel.core.cellstyle.ExcelCellStyleFactory;
import com.ihr360.excel.core.metaData.CellComment;
import com.ihr360.excel.core.metaData.ExportParams;
import com.ihr360.excel.util.Ihr360ExcelExportUtil;
import com.ihr360.excel.util.date.Ihr360ExcelDateFormatUtil;
import org.apache.commons.collections.map.HashedMap;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

public class TestExportBean {

    @Test
    public void exportXls() throws IOException {
        /**
         * 用排序的Map且key的顺序应与ExcelCell注解的index对应
         * 导出的数据顺序
         */
        Map<String, String> headerMap = new LinkedHashMap<>();
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_NAME, "姓名");
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_YEAR, "年");
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_MONTH, "月");
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_SALARY, "薪资");
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_TAX, "税额");
        headerMap.put(Payroll4ExcelVo.EXCEL_TEST_PAYTIME, "支付时间");

        ExcelCellStyle requiredStyle = ExcelCellStyleFactory.createRequiredHeaderCellStyle();




        Map<String, ExcelCellStyle> headerStyleMap = new HashedMap();
        headerStyleMap.put(Payroll4ExcelVo.EXCEL_TEST_NAME, requiredStyle);

        Collection<Payroll4ExcelVo> dataset = new ArrayList<>();
        dataset.add(new Payroll4ExcelVo("张三", 2017L, 12, 1234.0, 10.34, new Date()));
        dataset.add(new Payroll4ExcelVo("李四", 2017L, 10, 1345.0, 20.56, new Date()));
        dataset.add(new Payroll4ExcelVo("李四", 2017L, 11, null, 20.56, new Date()));
        File f = new File("exportBean.xls");
        OutputStream out = new FileOutputStream(f);

        //下拉列
//        Map<String, List<String>> dropDownsMap = new LinkedHashMap<>();
//        List<String> monthDropList = new ArrayList<>();
//        monthDropList.add("一月");
//        monthDropList.add("二月");
//        monthDropList.add("三月");
//        monthDropList.add("四月");
//        monthDropList.add("五月");
//        dropDownsMap.put(Payroll4ExcelVo.EXCEL_TEST_MONTH, monthDropList);

        //日期列输出格式
        Map<String, String> datePattern = new HashMap<>();
        datePattern.put(Payroll4ExcelVo.EXCEL_TEST_PAYTIME, Ihr360ExcelDateFormatUtil.PATTERN_ISO_ON_DATE);

        Map<String, CellComment> headerCommentMap = new HashMap<>();
        CellComment cellComment = CellComment.createCellComment(new int[]{255, 125, 1023, 150, 0, 0, 2, 2}, null, "这是姓名的备注", false);
        headerCommentMap.put(Payroll4ExcelVo.EXCEL_TEST_NAME, cellComment);


        ExportParams<Payroll4ExcelVo> exportParams = new ExportParams<>();
        exportParams.setHeaderMap(headerMap);
        exportParams.setRowDatas(dataset);
        exportParams.setHeaderStyleMap(headerStyleMap);
//        exportParams.setDropDownsMap(dropDownsMap);
        exportParams.setDatePatternMap(datePattern);
        exportParams.setHeaderCommentMap(headerCommentMap);

        Ihr360ExcelExportUtil.exportExcel(exportParams, out);

        out.close();
    }
}
