package com.ihr360.excel.handler;

import java.util.Collection;

/**
 * Excel处理接口
 *
 * @author richey
 */
public interface Ihr360ImportExcelProcessor<T> {

    <T> void doProcess();

    boolean isNext();

    boolean isBack() ;

    Collection<T> getResult();

    int getOrder();

}
