package com.dslengine.excel.builder.data.service


import groovy.transform.CompileStatic

@CompileStatic
abstract class QueryDataSource<C, I> implements DataSource<I> {

    private C c

    private void setDoQueryCondition(C c) {
        this.c = c
    }

    private void getDoQueryCondition(C c) {
        this.c = c
    }

    List<I> extractData(Long currentLoopTimes) {
        extractData(c, currentLoopTimes)
    }

    Long getMaxLoopTimes() {
        return null
    }

    abstract List<I> extractData(C c, Long currentLoopTimes)

}
