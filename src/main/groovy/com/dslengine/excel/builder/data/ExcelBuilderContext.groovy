package com.dslengine.excel.builder.data

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j

@Slf4j
@CompileStatic
final class ExcelBuilderContext {

    private static ThreadLocal<DataContext> dataContextHolder = new ThreadLocal()

    def static initDataContext(DataContext dataContext) {
        synchronized (dataContextHolder) {
            DataContext existedDataSource = dataContextHolder.get()
            if (existedDataSource) {

            }
            dataContextHolder.set(dataContext)
        }
    }

    def static getDataContext() {
        def dataContext
        synchronized (dataContextHolder) {
            DataContext existedDataSource = dataContextHolder.get()
            dataContext = existedDataSource
        }
        dataContext
    }

    def static releaseDataContext() {
        synchronized (dataContextHolder) {
            DataContext existedDataSource = dataContextHolder.get()
            if (existedDataSource) {

            }
            existedDataSource = null
            dataContextHolder.set(existedDataSource)
        }
    }

}
