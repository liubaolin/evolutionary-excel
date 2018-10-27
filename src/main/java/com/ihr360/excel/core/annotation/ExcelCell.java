package com.ihr360.excel.core.annotation;


import com.ihr360.excel.core.metaData.CellTypeMode;
import com.ihr360.excel.util.Ihr360ExcelImportUtil;
import org.apache.commons.lang3.StringUtils;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * The <code>ExcelCell</code><br>
 * <p>
 *     1.导入/导出时通过javaBean的方式，那么导入字段必须要加改注解，否则不会导入
 *     2.导出时如果用JavaBean的方式组织数据，则必须为index属性赋值，指定字段导出顺序
 * </p>
 * @see {@link Ihr360ExcelImportUtil#exportExcel}
 * @author richey.liu
 * @version 1.0 Created at 2017-12-17
 */

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCell {

    /**
     * 表头key，
     * 对应表头Map的key用于与表头对应
     *
     * @see {@link ExcelConfig#i18nStrategy()} 为 EXCEL_I18N_STRATEGY_PROPS时 可用于多语言词条key
     * @return header
     */
    String headerKey() default StringUtils.EMPTY;

    /**
     * 默认对应的表头名
     * 用作当导入必填字段找不到时，给出提示
     * 当allowNull为false时，此属性必要要有值
     * @return
     */
    String defaultHeaderName() default StringUtils.EMPTY;


    /**
     * 当值为null时要显示的值 default StringUtils.EMPTY
     * 
     * @return defaultValue
     */
    String defaultValue() default StringUtils.EMPTY;

    /**
     * Excel单元格验证模式
     * @return
     */
    CellTypeMode cellTypeMode() default CellTypeMode.LOOSE;

    /**
     * 是否弹性字段
     * flexibleField为true的字段用来放所有那些Excel表头没有对应固定字段的数据
     * 表注flexibleField为true的字段应该是有序的Map<表头，value>
     * @return
     */
    boolean flexibleField() default false;


    /**
     * 除Valid之外的校验
     */
    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.FIELD)
    @interface Valid {
        /**
         * 必须与in中String相符,目前仅支持String类型
         * 
         * @return e.g. {"key","i18nStrategy"}
         */
        String[] in() default {};

        /**
         * 是否允许为空,用于验证数据 default true
         * 
         * @return allowNull
         */
        boolean allowNull() default true;

        /**
         * Apply a "greater than" constraint to the named property
         * 
         * @return gt
         */
        double gt() default Double.NaN;

        /**
         * Apply a "less than" constraint to the named property
         * @return lt
         */
        double lt() default Double.NaN;

        /**
         * Apply a "greater than or equal" constraint to the named property
         * 
         * @return ge
         */
        double ge() default Double.NaN;

        /**
         * Apply a "less than or equal" constraint to the named property
         * 
         * @return le
         */
        double le() default Double.NaN;


    }
}
