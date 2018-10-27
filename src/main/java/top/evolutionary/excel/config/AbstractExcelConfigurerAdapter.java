package top.evolutionary.excel.config;

import top.evolutionary.excel.core.metaData.ImportParams;
import top.evolutionary.excel.commons.specification.ColumnSpecification;

import java.util.List;
import java.util.Map;

/**
 * @author richey
 */
public abstract class AbstractExcelConfigurerAdapter implements ExcelConfigurer {

    protected String rowNumKey = ExcelDefaultConfig.COMMON_SPECIFICATION_ROWNUM;
    protected boolean showRowNum = false;
    protected List<List<String>> atLeastOneOrIgnoreRow;
    protected List<List<String>> headerColumnJudge;
    protected boolean enableMatchStaff = true;
    protected List<String> staffNameAlias;
    protected List<String> mobileNoAlias;
    protected List<ColumnSpecification> columnSpecifications;

    protected ImportParams importParams;

    /**
     * 校验表头重复
     */
    protected boolean checkRepeatHeader = true;

    /**
     * 导入的类型 支持 javabean.class 或 Map.class
     */
    protected Class importType;

    /**
     * Map<String,List<String>>
     * importClazz为javaBean，且多语言策略为ExcelI18nStrategyType.EXCEL_I18N_STRATEGY_NONE时，
     * 必须提供importHeader，key用于同javabean的字段对应
     */
    protected Map<String, List<String>> importHeader;

    public AbstractExcelConfigurerAdapter() {

    }

    @Override
    public void setRowNumKey(String rowNumKey) {
        this.rowNumKey = rowNumKey;
    }

    @Override
    public void setShowRowNum(boolean showRowNum) {
        this.showRowNum = showRowNum;
    }


    @Override
    public void configureAtLeastOneOrIgnoreRow(List<List<String>> atLeastOneOrIgnoreRow) {
        this.atLeastOneOrIgnoreRow = atLeastOneOrIgnoreRow;
    }


    @Override
    public void configureHeaderColumnJudge(List<List<String>> headerColumnJudge) {
        this.headerColumnJudge = headerColumnJudge;
    }


    @Override
    public void enableMatchStaff(boolean enableMatchStaff) {
        this.enableMatchStaff = enableMatchStaff;
    }


    @Override
    public void configureStaffNameAlias(List<String> staffNameAlias) {
        this.staffNameAlias = staffNameAlias;
    }


    @Override
    public void configureMobileNoAlias(List<String> mobileNoAlias) {
        this.mobileNoAlias = mobileNoAlias;
    }

    @Override
    public void configureColumnSpecification(List<ColumnSpecification> columnSpecifications) {
        this.columnSpecifications = columnSpecifications;
    }

    @Override
    public void setImportType(Class type) {
        this.importType = type;
    }

    @Override
    public void configureImportHeader(Map<String, List<String>> importHeader) {
        this.importHeader = importHeader;
    }

    @Override
    public String getRowNumKey() {
        return rowNumKey;
    }

    @Override
    public boolean getShowRowNum() {
        return showRowNum;
    }

    @Override
    public List<List<String>> getAtLeastOneOrIgnoreRow() {
        return atLeastOneOrIgnoreRow;
    }

    @Override
    public List<List<String>> getHeaderColumnJudge() {
        return headerColumnJudge;
    }

    @Override
    public boolean getEnableMatchStaff() {
        return enableMatchStaff;
    }

    @Override
    public List<String> getStaffNameAlias() {
        return staffNameAlias;
    }

    @Override
    public List<String> getMobileNoAlias() {
        return mobileNoAlias;
    }

    @Override
    public List<ColumnSpecification> getColumnSpecification() {
        return columnSpecifications;
    }

    @Override
    public Class getImportType() {
        return importType;
    }

    @Override
    public Map<String, List<String>> getImportHeader() {
        return importHeader;
    }

    @Override
    public boolean isCheckRepeatHeader() {
        return checkRepeatHeader;
    }

    @Override
    public void setCheckRepeatHeader(boolean checkRepeatHeader) {
        this.checkRepeatHeader = checkRepeatHeader;
    }

    @Override
    public ImportParams getImportParam() {
        return importParams;
    }

    protected void setImportParam(ImportParams importParams) {
        this.importParams = importParams;
    }

}
