package com.ihr360.excel.event;

import com.ihr360.excel.commons.context.Ihr360ImportExcelContext;

import java.util.Collection;

/**
 * @author richey
 */
public abstract class ImportResultEventListener<T> {

    /**
     * when finish read excel(get the result) trigger invoke function
     *
     * @param result 结果集合
     */
    public abstract void invoke(Collection<T> result);

    /**
     * if have something to do after get result
     *
     * @param excelContext
     */
    public abstract void doAfterGetResult(Ihr360ImportExcelContext<T> excelContext);

}
