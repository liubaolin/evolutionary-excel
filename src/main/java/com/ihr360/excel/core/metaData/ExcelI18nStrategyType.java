package com.ihr360.excel.core.metaData;

/**
 *  Excel多语言策略
 */
public enum ExcelI18nStrategyType {

    /**
     * 不提供
     */
    EXCEL_I18N_STRATEGY_NONE(0),
    /**
     * 通过配置文件的方式
     */
    EXCEL_I18N_STRATEGY_PROPS(1);

    private final int strantegy;

    ExcelI18nStrategyType(int strantegy) {
        this.strantegy = strantegy;
    }

}
