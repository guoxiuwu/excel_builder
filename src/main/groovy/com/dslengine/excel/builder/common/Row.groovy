package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.layout.CellRangeBorderStyleApplier
import com.dslengine.excel.builder.layout.RowCellRangeBorderStyleApplier
import groovy.transform.CompileStatic
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet

class Row extends CreatesCells {

    private final XSSFRow row

    private int cellIdx

    Row(XSSFRow row, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder) {
        super(row.sheet, defaultOptions, columnIndexes, styleBuilder)
        this.row = row
        this.cellIdx = 0
    }

    @Override
    protected XSSFCell nextCell() {
        XSSFCell cell = row.createCell(cellIdx)
        cellIdx++
        cell
    }

    /**
     * @see CreatesCells#skipCells
     */
    @Override
    void skipCells(int num) {
        cellIdx += num
    }

    /**
     * Skip to a previously defined column created by {@link CreatesCells#column}
     *
     * @param id The column identifier
     */
    void skipTo(Object id) {
        if (columnIndexes && columnIndexes.containsKey(id)) {
            cellIdx = columnIndexes[id]
        } else {
            throw new IllegalArgumentException("Column index not specified for $id")
        }
    }

    /**
     * @see CreatesCells#merge(Map style, Closure callable)
     */
    @Override
    void merge(final Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        performMerge(style, callable)
    }

    @Override
    protected CellRangeAddress getRange(int start, int end) {
        new CellRangeAddress(row.rowNum, row.rowNum, start, end)
    }

    @Override
    protected int getMergeIndex() {
        cellIdx
    }

    @Override
    protected CellRangeBorderStyleApplier getBorderStyleApplier(CellRangeAddress range, XSSFSheet sheet) {
        new RowCellRangeBorderStyleApplier(range, sheet)
    }

    /**
     * @see CreatesCells#merge(Closure callable)
     */
    @Override
    void merge(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Row) Closure callable) {
        merge(null, callable)
    }
}
