package top.evolutionary.excel.core.metaData;

import top.evolutionary.excel.core.cellstyle.ExcelCellStyle;
import top.evolutionary.excel.commons.specification.ExportCommonSpecification;
import top.evolutionary.excel.commons.specification.MergedRegionSpecification;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * @author Richey.liu
 */
public class ExportParams<T> implements Serializable {

    private static final long serialVersionUID = -6266718905066495487L;
    /**
     * 需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *                    javabean属性的数据类型有基本数据类型及String,Date,String[],Double[]
     */
    private Collection<T> rowDatas;

    /**
     * 表头数据Map<columnKey,columnName>
     * 注意：要用有序的 LinkedHashMap！！！！
     */
    private Map<String, String> headerMap;

    /**
     * 表头批注
     */
    private Map<String,CellComment> headerCommentMap;

    /**
     * 用于表头合并的Map
     * 注意：要用有序的 LinkedHashMap！！！！
     */
    private Map<String, String> mergedHeaderMap;


    /**
     * 表头样式Map<columnKey,style>
     */
    private Map<String, ExcelCellStyle> headerStyleMap;

    /**
     * 下拉列<columnKey,List<dropItem>>
     */
    private Map<String, List<String>> dropDownsMap;

    /**
     * 用于多sheet导出
     */
    private List<ExcelSheet<T>> sheets;

    /**
     * 导出日期的pattern
     */
    private Map<String,String> datePatternMap;


    private Map<String,Class> dataTypeMap;


    private ExportCommonSpecification exportCommonSpecification;

    /**
     * 单元格合并<rowNum,specification>
     */
    private List<MergedRegionSpecification> mergedRegionSpecifications;

    public List<MergedRegionSpecification> getMergedRegionSpecifications() {
        return mergedRegionSpecifications;
    }

    public void setMergedRegionSpecifications(List<MergedRegionSpecification> mergedRegionSpecifications) {
        this.mergedRegionSpecifications = mergedRegionSpecifications;
    }

    public Map<String, String> getDatePatternMap() {
        return datePatternMap;
    }

    public void setDatePatternMap(Map<String, String> datePatternMap) {
        this.datePatternMap = datePatternMap;
    }

    public List<ExcelSheet<T>> getSheets() {
        return sheets;
    }

    public void setSheets(List<ExcelSheet<T>> sheets) {
        this.sheets = sheets;
    }

    public Collection<T> getRowDatas() {
        return rowDatas;
    }

    public void setRowDatas(Collection<T> rowDatas) {
        this.rowDatas = rowDatas;
    }

    public Map<String, String> getHeaderMap() {
        return headerMap;
    }

    public void setHeaderMap(Map<String, String> headerMap) {
        this.headerMap = headerMap;
    }

    public Map<String, ExcelCellStyle> getHeaderStyleMap() {
        return headerStyleMap;
    }

    public void setHeaderStyleMap(Map<String, ExcelCellStyle> headerStyleMap) {
        this.headerStyleMap = headerStyleMap;
    }

    public Map<String, List<String>> getDropDownsMap() {
        return dropDownsMap;
    }

    public void setDropDownsMap(Map<String, List<String>> dropDownsMap) {
        this.dropDownsMap = dropDownsMap;
    }

    public Map<String, String> getMergedHeaderMap() {
        return mergedHeaderMap;
    }

    public void setMergedHeaderMap(Map<String, String> mergedHeaderMap) {
        this.mergedHeaderMap = mergedHeaderMap;
    }

    public Map<String, CellComment> getHeaderCommentMap() {
        return headerCommentMap;
    }

    public void setHeaderCommentMap(Map<String, CellComment> headerCommentMap) {
        this.headerCommentMap = headerCommentMap;
    }

    public Map<String, Class> getDataTypeMap() {
        return dataTypeMap;
    }

    public void setDataTypeMap(Map<String, Class> dataTypeMap) {
        this.dataTypeMap = dataTypeMap;
    }

    public ExportCommonSpecification getExportCommonSpecification() {
        return exportCommonSpecification;
    }

    public void setExportCommonSpecification(ExportCommonSpecification exportCommonSpecification) {
        this.exportCommonSpecification = exportCommonSpecification;
    }
}
