package com.dslengine

import com.dslengine.excel.builder.common.ExcelBuilder
import com.dslengine.excel.builder.data.DataContext
import com.dslengine.excel.builder.data.service.DataSource
import com.dslengine.excel.builder.data.service.QueryDataSource
import org.junit.Test

class MainTest {

    @Test
    void test() {
        String outputFilePath = "D:/Test.xlsx"
        URL templateFilePath = this.class.getClassLoader().getResource("Test.dsl")

        File templateFile = new File(templateFilePath.toURI())
        File outputFile = new File(outputFilePath)
        DataContext dataContext = new MockData()

        dataContext.registerDataSource("registeredDataSource1", new DataSource1())
        dataContext.registerDataSource("registeredDataSource2", new DataSource2())

        ExcelBuilder.buildExcel(templateFile, outputFile, dataContext)
    }

    static class MockData implements DataContext {
        String field1 = "field1"
        String field2 = "field2"
        String field3 = "field3"
        String field4 = "field4"
        String field5 = "field5"

        List selfHoldList = [[key: "value1"], [key: "value2"], [key: "value3"], [key: "value4"], [key: "value5"]]
    }


    static class DataSource1 extends QueryDataSource {

        @Override
        List extractData(Object o, Long currentLoopTimes) {
            if (currentLoopTimes < 3) {
                return [1, 2, 3, 4, 5].collect {
                    [id: "Source1_LoopTime_${currentLoopTimes}_" + it]
                }
            }
            return []
        }
    }

    static class DataSource2 implements DataSource {

        @Override
        List extractData(Long currentLoopTimes) {
            return [1, 2, 3, 4, 5].collect {
                [id: "Source2_LoopTime_${currentLoopTimes}_" + it]
            }
        }

        @Override
        Long getMaxLoopTimes() {
            return 2L
        }
    }
}
