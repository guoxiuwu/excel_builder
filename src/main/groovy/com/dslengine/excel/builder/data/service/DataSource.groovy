package com.dslengine.excel.builder.data.service


interface DataSource<I> {

    // loop times start with 0
    List<I> extractData(Long currentLoopTimes)

    Long getMaxLoopTimes()

}
