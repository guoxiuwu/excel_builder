package com.dslengine.excel.builder.data

import com.dslengine.excel.builder.data.service.DataSource
import com.dslengine.excel.builder.exception.DataSourceRegistrationException
import groovy.transform.CompileStatic

@CompileStatic
trait DataContext {

    private final Map<String, DataSource> registeredDataSource = new HashMap()
    private def currentLoopDataItem

    def registerDataSource(String name, DataSource dataSource) {
        if (!name) {
            throw new DataSourceRegistrationException("the dataSource name must be provided")
        }
        DataSource existsDataSource = registeredDataSource.get(name)
        if (existsDataSource) {
            throw new DataSourceRegistrationException("the dataSource has existed with name '${name}'")
        }
        registeredDataSource.put(name, dataSource)
    }

    def findDataSource(String name) {
        if (!name) {
            throw new DataSourceRegistrationException("the dataSource name must be provided")
        }
        DataSource existsDataSource = registeredDataSource.get(name)
        if (!existsDataSource) {
            throw new DataSourceRegistrationException("the dataSource has not existed with name '${name}'")
        }
        existsDataSource
    }

    def cleanDataSource() {
        synchronized (registeredDataSource) {
            registeredDataSource.clear()
        }
    }

    def setDataItem(def item) {
        this.currentLoopDataItem = item
    }

    def getDataItem() {
        this.currentLoopDataItem
    }
}
