package com.ihr360.excel.util;

import com.ihr360.excel.handler.Ihr360ImportExcelProcessor;

import java.util.Collection;

/**
 * @author richey
 */
public abstract class AbstractIhr360ImportExcelProcessor<T> implements Ihr360ImportExcelProcessor<T> {

    protected int order = 0;

    protected boolean next = false;

    protected boolean back = false;

    private Collection<T> result;

    public AbstractIhr360ImportExcelProcessor(int order) {
        setOrder(order);
    }

    @Override
    public boolean isNext() {
        return next;
    }

    public void setNext(boolean next) {
        this.next = next;
    }

    @Override
    public boolean isBack() {
        return back;
    }

    public void setBack(boolean back) {
        this.back = back;
    }

    @Override
    public Collection<T> getResult() {
        return result;
    }

    public void setResult(Collection<T> result) {
        this.result = result;
    }

    @Override
    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }
}
