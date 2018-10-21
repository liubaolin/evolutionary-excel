package com.ihr360.excel.metaData;

/**
 * Excel导入列校验模式
 *
 */
public enum CellTypeMode {

    /**
     * 严格模式下，Excel导入列格式必须跟class相应字段的格式对应
     * 暂不支持，存在问题  Excel中设置格式跟poi读取格式有出入
     * eg：当文本类型在Excel直接改格式为数值类型时，poi读出来仍然时文本类型
     */
//    STRICT,
    /**
     * 松散模式下，class相关字段可能对应多种格式的Excel
     * eg：String类型的字段可以对应String，Number，Date 等类型的Excel单元格
     */
    LOOSE

}
