package com.ihr360.excel;

import com.ihr360.excel.config.Ihr360DefaultAbstractExcelConfiguration;
import org.junit.Test;

/**
 * @author richey
 */
public class TestApp {


    @Test
    public void test() {
        Ihr360DefaultAbstractExcelConfiguration configuration  = Ihr360DefaultAbstractExcelConfiguration.builder()
                .build();
    }

}
