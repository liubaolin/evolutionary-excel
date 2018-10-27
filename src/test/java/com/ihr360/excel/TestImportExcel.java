/**
 * @author SargerasWang
 */
package com.ihr360.excel;

import com.google.common.collect.Lists;
import com.ihr360.excel.config.Ihr360DefaultExcelConfiguration;
import com.ihr360.excel.config.ExcelDefaultConfig;
import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;
import com.ihr360.excel.commons.logs.ExcelLogType;
import com.ihr360.excel.commons.logs.ExcelLogs;
import com.ihr360.excel.commons.logs.ExcelRowLog;
import com.ihr360.excel.core.metaData.ImportParams;
import com.ihr360.excel.commons.specification.ColumnSpecification;
import com.ihr360.excel.commons.specification.CommonSpecification;
import com.ihr360.excel.util.Ihr360ExcelImportUtil;
import org.apache.commons.collections.CollectionUtils;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

/**
 * 测试导入
 */
public class TestImportExcel {

    /**
     * 导入输出javaBean类型数据
     *
     * @throws FileNotFoundException
     */
    @Test
    public void importBeanXls() throws FileNotFoundException {
        File f = new File("testImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);
        Ihr360ImportExcelContext ihr360ImportExcelContext = new Ihr360ImportExcelContext();

        ColumnSpecification nullBaleDateColumnSpe = ColumnSpecification.builder()
                .cellType(Date.class)
                .allowNull(true)
                .columns(Lists.newArrayList("弹性列5"))
                .build();

        Ihr360DefaultExcelConfiguration ihr360DefaultExcelConfiguration = Ihr360DefaultExcelConfiguration.builder()
                .importType(Payroll4ExcelVo.class)
                .columnSpecification(Lists.newArrayList(nullBaleDateColumnSpe))
                .build();


        Collection<Payroll4ExcelVo> excelDatas = Ihr360ExcelImportUtil.importExcel(ihr360DefaultExcelConfiguration, inputStream);

        ExcelLogs logs = ihr360ImportExcelContext.getLogs();

        //没有Excel基本错误日志
        assertTrue(!logs.hasExcelLogs());

        //有行日志信息
        assertTrue(logs.hasRowLogList());

        List<ExcelRowLog> excelRowLogs = logs.getRowLogList();
        Map<Integer, ExcelRowLog> numLogMap = excelRowLogs.stream().collect(Collectors.toMap(ExcelRowLog::getRowNum, Function.identity()));
        Map<String, Payroll4ExcelVo> nameEntityMap = excelDatas.stream().collect(Collectors.toMap(Payroll4ExcelVo::getName, Function.identity()));

        //用例测试
        assertNotNull(numLogMap.get(2));
        assertEquals(numLogMap.get(2).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(3));
        assertEquals(numLogMap.get(3).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(4));
        assertEquals(numLogMap.get(4).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(10));
        assertEquals(numLogMap.get(10).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(15));
        assertEquals(numLogMap.get(15).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(16));
        assertEquals(numLogMap.get(16).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(17));
        assertEquals(numLogMap.get(17).getExcelLogItems().size(), 1);
        assertNotNull(numLogMap.get(18));
        assertEquals(numLogMap.get(18).getExcelLogItems().size(), 1);

        assertEquals(numLogMap.get(2).getExcelLogItems().get(0).getLogType(), ExcelLogType.COLUMN_DATA_REQUIRED);
        assertEquals(numLogMap.get(3).getExcelLogItems().get(0).getLogType(), ExcelLogType.COLUMN_SCOPE_GE);
        assertEquals(numLogMap.get(4).getExcelLogItems().get(0).getLogType(), ExcelLogType.COLUMN_SCOPE_LE);
        //5/6/7/8/9行是文本日期格式导入，不应报错
        assertNull("文本类型日期格式 yyyy-MM-dd HH:mm:ss 解析失败", numLogMap.get(5));
        assertNull("文本类型日期格式 yyyy-MM-dd 解析失败", numLogMap.get(6));
        assertNull("文本类型日期格式 yyyy/MM/dd 解析失败", numLogMap.get(7));
        assertNull("文本类型日期格式 yyyy/MM/dd HH:mm:ss 解析失败", numLogMap.get(8));
        assertNull("文本类型日期格式 HH:mm:ss 解析失败", numLogMap.get(9));

        assertNotNull("文本类型日期格式 yyyy-MM-dd HH:mm:ss 解析到的数据为空", nameEntityMap.get("文本日期yyyy-MM-dd HH:mm:ss"));
        assertNotNull("文本类型日期格式 yyyy-MM-dd 解析到的数据为空", nameEntityMap.get("文本日期yyyy-MM-dd"));
        assertNotNull("文本类型日期格式 yyyy/MM/dd 解析到的数据为空", nameEntityMap.get("文本日期yyyy/MM/dd"));
        assertNotNull("文本类型日期格式 yyyy/MM/dd HH:mm:ss 解析到的数据为空", nameEntityMap.get("文本日期yyyy/MM/dd HH:mm:ss"));
        assertNotNull("文本类型日期格式 HH:mm:ss 解析到的数据为空", nameEntityMap.get("文本时间 HH:mm:ss"));
        //数据类型
        assertEquals(numLogMap.get(10).getExcelLogItems().get(0).getLogType(), ExcelLogType.COLUMN_FIELD_DATA_TYPE_ERR);

        assertNull(numLogMap.get(11));
        assertTrue("读取科学计数法数据不准确", nameEntityMap.get("读取科学计数法") != null && nameEntityMap.get("读取科学计数法").getYear().equals(26000000000L));
        assertNull(numLogMap.get(12));
        assertNotNull(nameEntityMap.get("超长小数取值不一致问题"));
        assertTrue("读取超长小数取值不一致", 1091.19649281798 == nameEntityMap.get("超长小数取值不一致问题").getSalary());
        assertNull(numLogMap.get(13));

        assertEquals(numLogMap.get(15).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(16).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(17).getExcelLogItems().get(0).getLogType(), ExcelLogType.BLANK_ROW);
        assertEquals(numLogMap.get(18).getExcelLogItems().get(0).getLogType(), ExcelLogType.COLUMN_FIELD_DATA_TYPE_ERR);


        printLogs(logs);

        for (Payroll4ExcelVo m : excelDatas) {
            System.out.println(m);
        }
    }

    private void printLogs(ExcelLogs logs) {
        if (logs.hasRowLogList()) {
            List<ExcelRowLog> errorLogList = logs.getRowLogList();
            System.out.println("Excel 数据行日志");
            for (ExcelRowLog errorLog : errorLogList) {
                if (CollectionUtils.isEmpty(errorLog.getExcelLogItems())) {
                    continue;
                }
                errorLog.getExcelLogItems().forEach(logItem -> {
                    System.out.println(errorLog.getRowNum() + " " + logItem);

                });
            }
        }
    }

    /**
     * 输出数据为Map<表头，数据>
     *
     * @throws FileNotFoundException
     */
    @Test
    public void importMapXls() throws FileNotFoundException {
        File f = new File("testImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);
        Ihr360ImportExcelContext ihr360ImportExcelContext = new Ihr360ImportExcelContext();


        Ihr360DefaultExcelConfiguration ihr360DefaultExcelConfiguration = Ihr360DefaultExcelConfiguration
                .builder()
                .importType(Map.class)
                .build();

        Collection excelDatas = Ihr360ExcelImportUtil.importExcel(ihr360DefaultExcelConfiguration, inputStream);
        List<ExcelRowLog> excelRowLogs = ihr360ImportExcelContext.getLogs().getRowLogList();
        Map<Integer, ExcelRowLog> numLogMap = excelRowLogs.stream().collect(Collectors.toMap(ExcelRowLog::getRowNum, Function.identity()));

        for (int i = 2; i <= 14; i++) {
            assertNull(numLogMap.get(2));
        }

        assertNotNull(numLogMap.get(15));
        assertNotNull(numLogMap.get(16));
        assertNotNull(numLogMap.get(17));

        assertEquals(numLogMap.get(15).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(16).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(17).getExcelLogItems().get(0).getLogType(), ExcelLogType.BLANK_ROW);

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

    /**
     * 测试导入数据，除指定列之外的列的格式
     *
     * @throws FileNotFoundException
     */
    @Test
    public void importSpecificationMapXls() throws FileNotFoundException {
        File f = new File("specificationImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);
        ExcelLogs logs = new ExcelLogs();

        ImportParams<Map> importParams = new ImportParams<>();
        importParams.setImportType(Map.class);
        ColumnSpecification ignoreColumnSpe = ColumnSpecification.createCellSpecification(String.class, true);
        ignoreColumnSpe.setIgnoreColumn(true);
        ignoreColumnSpe.addColumns("姓名");
        importParams.addColumnSpecification(ignoreColumnSpe);
        CommonSpecification commonSpecification = CommonSpecification.createCommonSpecification(true);
        importParams.setCommonSpecification(commonSpecification);

        List<List<String>> atLeastOneOrIgnoreRow = Lists.newArrayList();
        List<String> nameAilias = Lists.newArrayList("姓名", "名字");
        List<String> mobileAilias = Lists.newArrayList("手机", "手机号", "电话");
        atLeastOneOrIgnoreRow.add(nameAilias);
        atLeastOneOrIgnoreRow.add(mobileAilias);
        //至少有一列有值，否则忽略
        commonSpecification.setAtLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow);

        //表头判断，包含给定值的第一行认为是表头
        List<List<String>> headerJudgeList = Lists.newArrayList();
        headerJudgeList.add(nameAilias);
        headerJudgeList.add(mobileAilias);
        commonSpecification.setHeaderColumnJudge(headerJudgeList);

        List<Map> excelDatas = (List<Map>) Ihr360ExcelImportUtil.importExcel(importParams, inputStream);

        assertNotNull(excelDatas.get(0).get(ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM));
        assertEquals(excelDatas.get(0).get("支付时间"), "2018.01.29");
        assertEquals(excelDatas.get(1).get("支付时间"), "2018/01/29 01:02:03");
        assertEquals(excelDatas.get(2).get("支付时间"), "18-01-29");
        assertEquals(excelDatas.size(), 5);


        printLogs(logs);
        printExcelDatas(excelDatas);

    }

    /**
     * 测试导入数据，获取表头
     *
     * @throws FileNotFoundException
     */
    @Test
    public void importGetHeaderMap() throws FileNotFoundException {
        File f = new File("specificationImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);
        ExcelLogs logs = new ExcelLogs();

        ImportParams<Map> importParams = new ImportParams<>();
        importParams.setImportType(Map.class);

        Map<String, Integer> headerMap = Ihr360ExcelImportUtil.getHeaderTitleIndexMap();

        headerMap.forEach((k, v) -> {
            System.out.println("key=" + k + "" + "value=" + v);
        });

        logs.getRowLogList().forEach(excelRowLog -> System.out.println(excelRowLog.getExcelLogItems()));

    }


    /**
     * 测试导入数据，获取所有数据总数
     *
     * @throws FileNotFoundException
     */
    @Test
    public void importGetDataNum() throws FileNotFoundException {
        File f = new File("specificationImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);
        ExcelLogs logs = new ExcelLogs();

        ImportParams<Map> importParams = new ImportParams<>();
        importParams.setImportType(Map.class);

        Integer num = Ihr360ExcelImportUtil.countNorBlankOrHiddenRows();

        System.out.println(num);

    }


    @Test
    public void importXlsx() throws FileNotFoundException {
        File f = new File("testImportExcel.xls");
        InputStream inputStream = new FileInputStream(f);

        ExcelLogs logs = new ExcelLogs();
        ImportParams<Map> importParams = new ImportParams<>();
        importParams.setImportType(Map.class);


        Collection<Map> importExcel = Ihr360ExcelImportUtil.importExcel(importParams, inputStream);

        List<ExcelRowLog> excelRowLogs = logs.getRowLogList();

        Map<Integer, ExcelRowLog> numLogMap = excelRowLogs.stream().collect(Collectors.toMap(ExcelRowLog::getRowNum, Function.identity()));

        for (int i = 2; i <= 14; i++) {
            assertNull(numLogMap.get(2));
        }

        assertNotNull(numLogMap.get(15));
        assertNotNull(numLogMap.get(16));
        assertNotNull(numLogMap.get(17));
        assertEquals(numLogMap.get(15).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(16).getExcelLogItems().get(0).getLogType(), ExcelLogType.HIDDEN_ROW);
        assertEquals(numLogMap.get(17).getExcelLogItems().get(0).getLogType(), ExcelLogType.BLANK_ROW);

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }


}
