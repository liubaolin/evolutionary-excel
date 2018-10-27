package com.ihr360.excel.core.annotation;


import com.ihr360.excel.core.metaData.ExcelI18nStrategyType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel配置
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface ExcelConfig {

    /**
     * 多语言策略
     * @return
     */
    ExcelI18nStrategyType i18nStrategy() default ExcelI18nStrategyType.EXCEL_I18N_STRATEGY_NONE;

    /**
     * ClassPath路径下配置文件名称
     * @return
     */
    String propsFileName() default "";

    /**
     * 要全部删除的元素，默认删除所有换行和回车
     * @return
     */
    String globalRemovePattern() default "\r|\n";

    /**
     * 默认忽略隐藏行
     * @return
     */
    boolean ignoreHiddenRows() default true;

}
