package top.evolutionary.excel.commons.specification;

import top.evolutionary.excel.config.ExcelDefaultConfig;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * @author richey
 * @version 1.0.2, Created at 2018-03-27
 * 用于说明通用的规则，如是否返回rowNum
 * <p>
 * 目前只用于map结构的导入数据
 * 暂不支持javaBean自身的field，javaBean自身的field暂时只用注解配置
 */
public class CommonSpecification implements Serializable {

    private String rowNumKey = ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM;

    private boolean showRowNum = false;

    /**
     * 给定的列中至少有一列有值，否则忽略
     * 列以List集合的形式存在，可以有多个别名，优先级与顺序相同，存在多列时，根据优先级匹配判断
     */
    private List<List<String>> atLeastOneOrIgnoreRow;

    /**
     * 表头判断，如果包含所有给定的表头，则认为该行数据是表头行
     * 当前会取第一个表头行
     */
    private List<List<String>> headerColumnJudge;

    /**
     * 是否匹配员工
     * 默认开启，如果匹配上会根据手机号和姓名匹配员工id数据
     */
    private boolean enableMatchStaff = true;

    /**
     * 员工姓名的所有别名，按顺序，取第一个
     * 员工姓名 + 手机号 匹配唯一员工
     */
    private List<String> staffNameAlias;

    /**
     * 员工手机号的所有别名，按顺序，取第一个
     * 员工姓名 + 手机号 匹配唯一员工
     */
    private List<String> mobileNoAlias;


    /**
     * 模板表头
     */
    protected Map<String, Integer> templateHeaderIndexTitleMap;
    /**
     * 模板表头所在列
     */
    protected List<Integer> templateHeaderRowNums;
    /**
     * 模板数据开始行
     */
    protected Integer templateDataBeginRowNum;

    /**
     * 校验表头重复
     */
    protected boolean checkRepeatHeader = true;


    public CommonSpecification() {
    }

    public CommonSpecification(String rowNumKey, boolean showRowNum) {
        this.rowNumKey = rowNumKey;
        this.showRowNum = showRowNum;
    }

    public String getRowNumKey() {
        return rowNumKey;
    }

    public void setRowNumKey(String rowNumKey) {
        this.rowNumKey = rowNumKey;
    }

    public boolean isShowRowNum() {
        return showRowNum;
    }

    public void setShowRowNum(boolean showRowNum) {
        this.showRowNum = showRowNum;
    }

    @Deprecated
    public static CommonSpecification createCommonSpecification(boolean showRowNum) {
        return new CommonSpecification(ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM, showRowNum);
    }

    public List<List<String>> getAtLeastOneOrIgnoreRow() {
        return atLeastOneOrIgnoreRow;
    }

    public void setAtLeastOneOrIgnoreRow(List<List<String>> atLeastOneOrIgnoreRow) {
        this.atLeastOneOrIgnoreRow = atLeastOneOrIgnoreRow;
    }

    public List<List<String>> getHeaderColumnJudge() {
        return headerColumnJudge;
    }

    public void setHeaderColumnJudge(List<List<String>> headerColumnJudge) {
        this.headerColumnJudge = headerColumnJudge;
    }

    public List<String> getStaffNameAlias() {
        return staffNameAlias;
    }

    public void setStaffNameAlias(List<String> staffNameAlias) {
        this.staffNameAlias = staffNameAlias;
    }

    public List<String> getMobileNoAlias() {
        return mobileNoAlias;
    }

    public void setMobileNoAlias(List<String> mobileNoAlias) {
        this.mobileNoAlias = mobileNoAlias;
    }

    public boolean isEnableMatchStaff() {
        return enableMatchStaff;
    }

    public void setEnableMatchStaff(boolean enableMatchStaff) {
        this.enableMatchStaff = enableMatchStaff;
    }

    public Map<String, Integer> getTemplateHeaderIndexTitleMap() {
        return templateHeaderIndexTitleMap;
    }

    @Deprecated
    public void setTemplateHeaderIndexTitleMap(Map<String, Integer> templateHeaderIndexTitleMap) {
        this.templateHeaderIndexTitleMap = templateHeaderIndexTitleMap;
    }

    public List<Integer> getTemplateHeaderRowNums() {
        return templateHeaderRowNums;
    }

    @Deprecated
    public void setTemplateHeaderRowNums(List<Integer> templateHeaderRowNums) {
        this.templateHeaderRowNums = templateHeaderRowNums;
    }

    public Integer getTemplateDataBeginRowNum() {
        return templateDataBeginRowNum;
    }

    @Deprecated
    public void setTemplateDataBeginRowNum(Integer templateDataBeginRowNum) {
        this.templateDataBeginRowNum = templateDataBeginRowNum;
    }

    public boolean isCheckRepeatHeader() {
        return checkRepeatHeader;
    }

    public void setCheckRepeatHeader(boolean checkRepeatHeader) {
        this.checkRepeatHeader = checkRepeatHeader;
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

        public Builder templateHeaders(Map<String, Integer> templateHeaders) {
            this.templateHeaders = templateHeaders;
            return this;
        }

        public Builder templateHeaderRowNums(List<Integer> templateHeaderRowNums) {
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


        public CommonSpecification build() {
            CommonSpecification commonSpecification = new CommonSpecification();
            commonSpecification.setRowNumKey(rowNumKey);
            commonSpecification.setShowRowNum(showRowNum);
            commonSpecification.setAtLeastOneOrIgnoreRow(atLeastOneOrIgnoreRow);
            commonSpecification.setHeaderColumnJudge(headerColumnJudge);
            commonSpecification.setEnableMatchStaff(enableMatchStaff);
            commonSpecification.setStaffNameAlias(staffNameAlias);
            commonSpecification.setMobileNoAlias(mobileNoAlias);
            commonSpecification.setTemplateHeaderIndexTitleMap(templateHeaders);
            commonSpecification.setTemplateHeaderRowNums(templateHeaderRowNums);
            commonSpecification.setTemplateDataBeginRowNum(templateDataBeginRowNum);
            commonSpecification.setCheckRepeatHeader(checkRepeatHeader);
            return commonSpecification;
        }
    }


}
