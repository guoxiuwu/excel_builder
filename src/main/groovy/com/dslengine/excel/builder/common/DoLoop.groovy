package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.data.Base
import com.dslengine.excel.builder.data.service.DataSource
import com.dslengine.excel.builder.exception.DSLSyntaxException

trait DoLoop extends Base {

    def loop(String dataSourceName, Closure cl) {
        if (!dataSourceName) {
            throw new DSLSyntaxException("dataSourceName must be provided in 'Loop()'")
        }
        if (!cl) {
            throw new DSLSyntaxException("invalid loop defined,it should be like 'Loop('dataSourceName'){ ... }'")
        }
        DataSource dataSource = getDataContext().findDataSource(dataSourceName)

        Long maxLoopTimes = dataSource.getMaxLoopTimes()
        if (maxLoopTimes) {
            for (long i; i < maxLoopTimes; i++) {
                List items = dataSource.extractData(i)
                if (items && items.size() > 0) {
                    executeDataLoop(items, cl)
                }
            }
        } else {
            Long loopTimes = 0L
            while (true) {
                List items = dataSource.extractData(loopTimes)
                if (items && items.size() > 0) {
                    executeDataLoop(items, cl)
                } else {
                    break
                }
                loopTimes++
            }
        }
    }

    def loop(Collection loopDataList, Closure cl) {
        if (!cl) {
            throw new DSLSyntaxException("invalid loop defined,it should be like 'Loop('dataSourceName'){ ... }'")
        }
        executeDataLoop(loopDataList, cl)
    }


    private void executeDataLoop(List dataItems, Closure cl) {
        dataItems?.each { item ->
            getDataContext().setDataItem(item)
            // here should be replaced with parent's delegate object
            cl.setDelegate(this)
            cl.setResolveStrategy(Closure.DELEGATE_FIRST)
            cl.call()
        }
    }
}
