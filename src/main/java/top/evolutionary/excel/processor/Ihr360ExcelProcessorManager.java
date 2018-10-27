package top.evolutionary.excel.processor;

import com.google.common.collect.Lists;
import org.apache.commons.collections.CollectionUtils;

import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * @author richey
 */
public class Ihr360ExcelProcessorManager<T> {

    private List<Ihr360ImportExcelProcessor<T>> processors;

    public Ihr360ExcelProcessorManager() {
        processors = Lists.newArrayListWithCapacity(2);
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
