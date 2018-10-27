package com.ihr360.excel.config;

import com.ihr360.excel.core.metaData.ImportParams;
import com.ihr360.excel.commons.specification.ColumnSpecification;
import com.ihr360.excel.commons.specification.CommonSpecification;

import java.util.List;
import java.util.Map;

/**
 * 默认导入配置
 *
 * @author richey
 */
public class Ihr360DefaultExcelConfiguration extends AbstractExcelConfigurerAdapter {

    private Ihr360DefaultExcelConfiguration() {
        super();
    }

    public static Ihr360DefaultExcelConfiguration.Builder builder() {
        return new Ihr360DefaultExcelConfiguration.Builder();
    }


    public static class Builder {

        private String rowNumKey = ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM;
        private boolean showRowNum = false;
        private List<List<String>> atLeastOneOrIgnoreRow;
        private List<List<String>> headerColumnJudge;
        private boolean enableMatchStaff = true;
        private List<String> staffNameAlias;
        private List<String> mobileNoAlias;
        private List<ColumnSpecification> columnSpecifications;
        private Class importType;
        private Map<String, List<String>> importHeader;

        public Builder rowNumKey(String rowNumKey) {
            this.rowNumKey = rowNumKey;
            return this;
        }

        public Builder showRowNum(boolean showRowNum) {
            this.showRowNum = showRowNum;
            return this;
        }

        public Builder atLeastOneOrIgnoreRow(List<List<String>> atLeastOneOrIgnoreRow) {
            this.atLeastOneOrIgnoreRow = atLeastOneOrIgnoreRow;
            return this;
        }


        public Builder headerColumnJudge(List<List<String>> headerColumnJudge) {
            this.headerColumnJudge = headerColumnJudge;
            return this;
        }


        public Builder enableMatchStaff(boolean enableMatchStaff) {
            this.enableMatchStaff = enableMatchStaff;
            return this;
        }


        public Builder staffNameAlias(List<String> staffNameAlias) {
            this.staffNameAlias = staffNameAlias;
            return this;
        }


        public Builder mobileNoAlias(List<String> mobileNoAlias) {
            this.mobileNoAlias = mobileNoAlias;
            return this;
        }

        public Builder columnSpecification(List<ColumnSpecification> columnSpecifications) {
            this.columnSpecifications = columnSpecifications;
            return this;
        }


        public Builder importType(Class type) {
            this.importType = type;
            return this;
        }

        public Builder importHeader(Map<String, List<String>> importHeader) {
            this.importHeader = importHeader;
            return this;
        }


        public Ihr360DefaultExcelConfiguration build() {
            Ihr360DefaultExcelConfiguration ihr360DefaultExcelConfiguration = new Ihr360DefaultExcelConfiguration();

            CommonSpecification commonSpecification = CommonSpecification.builder()
                    .rowNumKey(rowNumKey)
                    .showRowNum(showRowNum)
                    .atLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow)
                    .headerColumnJudge(headerColumnJudge)
                    .enableMatchStaff(enableMatchStaff)
                    .staffNameAlias(staffNameAlias)
                    .mobileNoAlias(mobileNoAlias)
                    .build();

            ImportParams importParams = new ImportParams();
            importParams.setCommonSpecification(commonSpecification);
            importParams.setColumnSpecifications(columnSpecifications);
            importParams.setImportType(importType);
            importParams.setImportHeader(importHeader);
            ihr360DefaultExcelConfiguration.setImportParam(importParams);

            String rowNumKey = commonSpecification.getRowNumKey();
            boolean showRowNum = commonSpecification.isShowRowNum();
            List<List<String>> atLeastOneOrIgnoreRow = commonSpecification.getAtLeastOneOrIgnoreRow();
            List<List<String>> headerColumnJudge = commonSpecification.getHeaderColumnJudge();
            boolean enableMatchStaff = commonSpecification.isEnableMatchStaff();
            List<String> staffNameAlias = commonSpecification.getStaffNameAlias();
            List<String> mobileNoAlias = commonSpecification.getMobileNoAlias();
            ihr360DefaultExcelConfiguration.setRowNumKey(rowNumKey);
            ihr360DefaultExcelConfiguration.setShowRowNum(showRowNum);
            ihr360DefaultExcelConfiguration.configureAtLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow);
            ihr360DefaultExcelConfiguration.configureHeaderColumnJudge(headerColumnJudge);
            ihr360DefaultExcelConfiguration.enableMatchStaff(enableMatchStaff);
            ihr360DefaultExcelConfiguration.configureStaffNameAlias(staffNameAlias);
            ihr360DefaultExcelConfiguration.configureMobileNoAlias(mobileNoAlias);
            ihr360DefaultExcelConfiguration.configureColumnSpecification(columnSpecifications);
            ihr360DefaultExcelConfiguration.setImportType(importType);
            ihr360DefaultExcelConfiguration.configureImportHeader(importHeader);

            return ihr360DefaultExcelConfiguration;
        }

    }


}
