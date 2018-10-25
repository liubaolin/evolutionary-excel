package com.ihr360.excel.handler;

import com.ihr360.excel.util.Ihr360ExcelHeaderProcessor;
import com.ihr360.excel.util.Ihr360ExcelRowDataProcessor;
import org.apache.commons.collections.CollectionUtils;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * @author richey
 */
public class Ihr360ExcelHandlerManager<T> {

    private List<Ihr360ImportExcelProcessor<T>> processors = Collections.emptyList();

    public Ihr360ExcelHandlerManager() {
        processors.add(new Ihr360ExcelHeaderProcessor<T>(0));
        processors.add(new Ihr360ExcelRowDataProcessor<T>(1));
    }

    public List<Ihr360ImportExcelProcessor<T>> getProcessors() {
        if (CollectionUtils.isEmpty(processors)) {
            return Collections.EMPTY_LIST;
        }
        processors.sort(Comparator.comparing(Ihr360ImportExcelProcessor::getOrder));
        return processors;
    }

}
