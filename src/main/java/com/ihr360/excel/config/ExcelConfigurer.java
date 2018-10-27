package com.ihr360.excel.config;


import com.ihr360.excel.core.metaData.ImportParams;
import com.ihr360.excel.commons.specification.ColumnSpecification;

import java.util.List;
import java.util.Map;

public interface ExcelConfigurer {

    void setRowNumKey(String rowNumKey);

    void setShowRowNum(boolean showRowNum);

    /**
     * 给定的列中至少有一列有值，否则忽略
     * 列以List集合的形式存在，可以有多个别名，优先级与顺序相同，存在多列时，根据优先级匹配判断
     */
    void configureAtLeastOneOrIgnoreRow(List<List<String>> atLeastOneOrIgnoreRow);

    /**
     * 表头判断，如果包含所有给定的表头，则认为该行数据是表头行
     * 当前会取第一个表头行
     */
    void configureHeaderColumnJudge(List<List<String>> headerColumnJudge);

    /**
     * 是否匹配员工
     * 默认开启，如果匹配上会根据手机号和姓名匹配员工id数据
     */
    void enableMatchStaff(boolean enableMatchStaff);

    /**
     * 员工姓名的所有别名，按顺序，取第一个
     * 员工姓名 + 手机号 匹配唯一员工
     */
    void configureStaffNameAlias(List<String> staffNameAlias);

    /**
     * 员工手机号的所有别名，按顺序，取第一个
     * 员工姓名 + 手机号 匹配唯一员工
     */
    void configureMobileNoAlias(List<String> mobileNoAlias);


    void configureColumnSpecification(List<ColumnSpecification> columnSpecifications);

    void setImportType(Class type);

    void configureImportHeader(Map<String, List<String>> importHeader);

    String getRowNumKey();

    boolean getShowRowNum();

    List<List<String>> getAtLeastOneOrIgnoreRow();

    List<List<String>> getHeaderColumnJudge();

    boolean getEnableMatchStaff();

    List<String> getStaffNameAlias();

    List<String> getMobileNoAlias();

    List<ColumnSpecification> getColumnSpecification();

    Class getImportType();

    Map<String, List<String>> getImportHeader();

    boolean isCheckRepeatHeader();

    void setCheckRepeatHeader(boolean checkRepeatHeader);

    ImportParams getImportParam();

}
