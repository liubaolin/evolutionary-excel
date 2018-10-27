package top.evolutionary.excel.core.metaData;

import top.evolutionary.excel.commons.specification.ColumnSpecification;
import top.evolutionary.excel.commons.specification.CommonSpecification;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author richey
 * @param <T>
 */
public class ImportParams<T>{

    /**
     * 导入的类型 支持 javabean.class 或 Map.class
     */
    private Class<T> importType;

    /**
     * Map<String,List<String>>
     * importClazz为javaBean，且多语言策略为ExcelI18nStrategyType.EXCEL_I18N_STRATEGY_NONE时，
     * 必须提供importHeader，key用于同javabean的字段对应
     */
    private Map<String, List<String>> importHeader;

    /**
     * 单元格规则
     */
    private List<ColumnSpecification> columnSpecifications = new ArrayList<>();

    private CommonSpecification commonSpecification;

    public ImportParams() {
    }


    public void setImportType(Class<T> importType) {
        this.importType = importType;
    }

    public Class<T> getImportType() {
        return importType;
    }

    public Map<String, List<String>> getImportHeader() {
        return importHeader;
    }

    public void setImportHeader(Map<String, List<String>> importHeader) {
        this.importHeader = importHeader;
    }

    public List<ColumnSpecification> getColumnSpecifications() {
        return columnSpecifications;
    }

    public void setColumnSpecifications(List<ColumnSpecification> columnSpecifications) {
        this.columnSpecifications = columnSpecifications;
    }

    public CommonSpecification getCommonSpecification() {
        return commonSpecification;
    }

    public void setCommonSpecification(CommonSpecification commonSpecification) {
        this.commonSpecification = commonSpecification;
    }

    public void addColumnSpecification(ColumnSpecification columnSpecification) {
        this.columnSpecifications.add(columnSpecification);
    }
}
