package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.layout.CellRangeBorderStyleApplier
import com.dslengine.excel.builder.layout.ColumnCellRangeBorderStyleApplier
import groovy.transform.CompileStatic
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy


@CompileStatic
class Column extends CreatesCells {

    private int columnIdx
    private int rowIdx

    Column(XSSFSheet sheet, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder, int columnIdx, int rowIdx) {
        super(sheet, defaultOptions, columnIndexes, styleBuilder)
        this.columnIdx = columnIdx
        this.rowIdx = rowIdx
    }

    @Override
    protected XSSFCell nextCell() {
        XSSFRow row = sheet.getRow(rowIdx)
        if (row == null) {
            row = sheet.createRow(rowIdx)
        }
        XSSFCell cell = row.getCell(columnIdx, MissingCellPolicy.CREATE_NULL_AS_BLANK)
        rowIdx++
        cell
    }

    /**
     * @see CreatesCells#skipCells
     */
    void skipCells(int num) {
        rowIdx += num
    }

    /**
     * @see CreatesCells#merge(Map style, Closure callable)
     */
    @Override
    void merge(Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        performMerge(style, callable)
    }

    @Override
    protected CellRangeAddress getRange(int start, int end) {
        new CellRangeAddress(start, end, columnIdx, columnIdx)
    }

    @Override
    protected int getMergeIndex() {
        rowIdx
    }

    @Override
    protected CellRangeBorderStyleApplier getBorderStyleApplier(CellRangeAddress range, XSSFSheet sheet) {
        new ColumnCellRangeBorderStyleApplier(range, sheet)
    }

    /**
     * @see CreatesCells#merge(Closure callable)
     */
    @Override
    void merge(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Column) Closure callable) {
        merge(null, callable)
    }
}
