package top.evolutionary.excel.config;

import top.evolutionary.excel.core.metaData.ImportParams;
import top.evolutionary.excel.commons.specification.ColumnSpecification;
import top.evolutionary.excel.commons.specification.CommonSpecification;

import java.util.List;
import java.util.Map;

/**
 * 应用于模板的导入配置
 */
public class Ihr360TemplateExcelConfiguration extends AbstractExcelConfigurerAdapter {


    public Ihr360TemplateExcelConfiguration() {
        super();
    }

    /**
     * 模板表头 表头－＞colIndex
     */
    protected Map<String, Integer> templateHeaders;
    /**
     * 模板表头所在行,如果存在多行，按顺序后面相同索引的列覆盖前面的
     */
    protected List<Integer> templateHeaderRowNums;
    /**
     * 模板数据开始行
     */
    protected Integer templateDataBeginRowNum;


    public void configerTemplateHeaders(Map<String, Integer> templateHeaders) {
        this.templateHeaders = templateHeaders;
    }

    public void setTemplateDataBeginRowNum(Integer templateDataBeginRowNum) {
        this.templateDataBeginRowNum = templateDataBeginRowNum;
    }

    public Map<String, Integer> getTemplateHeaders() {
        return this.templateHeaders;
    }


    public Integer getTemplateDataBeginRowNum() {
        return this.templateDataBeginRowNum;
    }

    public void setTemplateHeaders(Map<String, Integer> templateHeaders) {
        this.templateHeaders = templateHeaders;
    }

    public List<Integer> getTemplateHeaderRowNums() {
        return templateHeaderRowNums;
    }

    public void setTemplateHeaderRowNums(List<Integer> templateHeaderRowNums) {
        this.templateHeaderRowNums = templateHeaderRowNums;
    }

    public static Builder builder() {
        return new Builder();
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
        private Map<String, Integer> templateHeaders;
        private List<Integer> templateHeaderRowNums;
        private Integer templateDataBeginRowNum;
        private boolean checkRepeatHeader = true;


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

        public Builder templateHeaders(Map<String, Integer> templateHeaders) {
            this.templateHeaders = templateHeaders;
            return this;
        }

        public Builder templateHeaderRowNum(List<Integer> templateHeaderRowNums) {
            this.templateHeaderRowNums = templateHeaderRowNums;
            return this;
        }

        public Builder templateDataBeginRowNum(Integer templateDataBeginRowNum) {
            this.templateDataBeginRowNum = templateDataBeginRowNum;
            return this;
        }

        public Builder checkRepeatHeader(boolean checkRepeatHeader) {
            this.checkRepeatHeader = checkRepeatHeader;
            return this;
        }

        public Ihr360TemplateExcelConfiguration build() {

            CommonSpecification commonSpecification = CommonSpecification.builder()
                    .rowNumKey(rowNumKey)
                    .showRowNum(showRowNum)
                    .atLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow)
                    .headerColumnJudge(headerColumnJudge)
                    .enableMatchStaff(enableMatchStaff)
                    .staffNameAlias(staffNameAlias)
                    .mobileNoAlias(mobileNoAlias)
                    .checkRepeatHeader(checkRepeatHeader)
                    .build();

            Ihr360TemplateExcelConfiguration ihr360TemplateExcelConfiguration = new Ihr360TemplateExcelConfiguration();

            ImportParams importParams = new ImportParams();
            importParams.setCommonSpecification(commonSpecification);
            importParams.setColumnSpecifications(columnSpecifications);
            importParams.setImportType(importType);
            importParams.setImportHeader(importHeader);
            ihr360TemplateExcelConfiguration.setImportParam(importParams);

            ihr360TemplateExcelConfiguration.setRowNumKey(rowNumKey);
            ihr360TemplateExcelConfiguration.setShowRowNum(showRowNum);
            ihr360TemplateExcelConfiguration.configureAtLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow);
            ihr360TemplateExcelConfiguration.configureHeaderColumnJudge(headerColumnJudge);
            ihr360TemplateExcelConfiguration.enableMatchStaff(enableMatchStaff);
            ihr360TemplateExcelConfiguration.configureStaffNameAlias(staffNameAlias);
            ihr360TemplateExcelConfiguration.configureMobileNoAlias(mobileNoAlias);
            ihr360TemplateExcelConfiguration.configureColumnSpecification(columnSpecifications);
            ihr360TemplateExcelConfiguration.setImportType(importType);
            ihr360TemplateExcelConfiguration.configureImportHeader(importHeader);
            ihr360TemplateExcelConfiguration.configerTemplateHeaders(templateHeaders);
            ihr360TemplateExcelConfiguration.setTemplateHeaderRowNums(templateHeaderRowNums);
            ihr360TemplateExcelConfiguration.setTemplateDataBeginRowNum(templateDataBeginRowNum);
            ihr360TemplateExcelConfiguration.setCheckRepeatHeader(checkRepeatHeader);
            return ihr360TemplateExcelConfiguration;
        }

    }


}
