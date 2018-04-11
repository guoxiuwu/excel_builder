package com.dslengine.excel.builder.data

import com.dslengine.excel.builder.exception.DataContextException

trait Base {

    DataContext getDataContext() {
        DataContext dataContext = ExcelBuilderContext.getDataContext()
        if (!dataContext) {
            throw new DataContextException("dataContext has not been initialized")
        }
        dataContext
    }

    def getDataItem() {
        def dataItem = getDataContext().getDataItem()
        if (!dataItem) {
            throw new DataContextException("Current DataItem has not been initialized")
        }
        dataItem
    }
}
