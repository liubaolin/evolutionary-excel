package com.ihr360.excel.constants;

import com.ihr360.excel.util.ExcelDateFormatUtil;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelDefaultConfig {

    //日期类型默认输出格式
    public static final String DEFAULT_OUTPUT_DATE_PATTERN = ExcelDateFormatUtil.PATTERN_DEFAULT_ON_SECOND;

    //注解的排序字段
    public static final String SORT_ANNO_PROPS = "index";

    public static final short DEFAULT_ROW_HEADER_HEIGHT_INPOINT = 20;

    //行号在Map中的Key
    public static final String COMMON_SPECIFICATION_ROWNUM = "rowNum";


    /**
     * 用来验证excel与Vo中的类型是否一致 <br>
     * Map<栏位类型,只能是哪些Cell类型>
     * 松散模式
     *
     * poi-11版本问题
     */
//    public static Map<Class<?>, CellType[]> looseValidateMap = new HashMap<>();
    public static Map<Class<?>, Integer[]> looseValidateMap = new HashMap<>();

    /**
     * 严格模式
     * poi-11版本问题
     */
//    public static Map<Class<?>, CellType[]> strickValidateMap = new HashMap<>();
    public static Map<Class<?>, int[]> strickValidateMap = new HashMap<>();

    static {
        //暂不支持数组类型
//        looseValidateMap.put(String[].class, new CellType[]{CellType.STRING});
//        looseValidateMap.put(Double[].class, new CellType[]{CellType.NUMERIC});

        /*looseValidateMap.put(String.class, new CellType[]{CellType.STRING, CellType.BLANK,CellType.NUMERIC});
        looseValidateMap.put(Double.class, new CellType[]{CellType.NUMERIC,CellType.STRING});
        looseValidateMap.put(Date.class, new CellType[]{CellType.NUMERIC, CellType.STRING});
        looseValidateMap.put(Integer.class, new CellType[]{CellType.NUMERIC,CellType.STRING});
        looseValidateMap.put(Float.class, new CellType[]{CellType.NUMERIC,CellType.STRING});
        looseValidateMap.put(Long.class, new CellType[]{CellType.NUMERIC,CellType.STRING});
        looseValidateMap.put(Boolean.class, new CellType[]{CellType.BOOLEAN});*/

        //poi-11版本问题
        looseValidateMap.put(String.class, new Integer[]{Cell.CELL_TYPE_STRING, Cell.CELL_TYPE_BLANK,Cell.CELL_TYPE_NUMERIC});
        looseValidateMap.put(Double.class, new Integer[]{Cell.CELL_TYPE_NUMERIC,Cell.CELL_TYPE_STRING});
        looseValidateMap.put(Date.class, new Integer[]{Cell.CELL_TYPE_NUMERIC, Cell.CELL_TYPE_STRING});
        looseValidateMap.put(Integer.class, new Integer[]{Cell.CELL_TYPE_NUMERIC,Cell.CELL_TYPE_STRING});
        looseValidateMap.put(Float.class, new Integer[]{Cell.CELL_TYPE_NUMERIC,Cell.CELL_TYPE_STRING});
        looseValidateMap.put(Long.class, new Integer[]{Cell.CELL_TYPE_NUMERIC,Cell.CELL_TYPE_STRING});
        looseValidateMap.put(Boolean.class, new Integer[]{Cell.CELL_TYPE_BOOLEAN});


      /*  strickValidateMap.put(String.class, new CellType[]{CellType.STRING, CellType.BLANK});
        strickValidateMap.put(Double.class, new CellType[]{CellType.NUMERIC});
        strickValidateMap.put(Date.class, new CellType[]{CellType.NUMERIC});
        strickValidateMap.put(Integer.class, new CellType[]{CellType.NUMERIC});
        strickValidateMap.put(Float.class, new CellType[]{CellType.NUMERIC});
        strickValidateMap.put(Long.class, new CellType[]{CellType.NUMERIC});
        strickValidateMap.put(Boolean.class, new CellType[]{CellType.BOOLEAN});*/
    }





}
