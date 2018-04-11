package com.dslengine.excel.builder.layout

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import groovy.transform.CompileStatic
import groovy.transform.TupleConstructor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide

@TupleConstructor
abstract class CellRangeBorderStyleApplier implements BorderStyleApplier {

    CellRangeAddress range
    Sheet sheet

    protected XSSFCellStyle leftTop
    protected XSSFCellStyle middle
    protected XSSFCellStyle bottomRight

    CellRangeBorderStyleApplier(CellRangeAddress range, XSSFSheet sheet) {
        this.range = range
        this.sheet = sheet
        leftTop = sheet.workbook.createCellStyle()
        if (range.numberOfCells > 2) {
            middle = sheet.workbook.createCellStyle()
        }
        bottomRight = sheet.workbook.createCellStyle()

        Row row = CellUtil.getRow(range.firstRow, sheet)
        leftTop.cloneStyleFrom(CellUtil.getCell(row, range.firstColumn).cellStyle)
        middle?.cloneStyleFrom(leftTop)
        bottomRight.cloneStyleFrom(leftTop)
    }

    @Override
    abstract void applyStyle(BorderSide side, BorderStyle style)

    @Override
    void applyStyle(BorderStyle style) {
        BorderSide.values().each { BorderSide side ->
            applyStyle(side, style)
        }
    }

    @Override
    abstract void applyColor(BorderSide side, XSSFColor color)

    @Override
    void applyColor(XSSFColor color) {
        BorderSide.values().each { BorderSide side ->
            applyColor(side, color)
        }
    }

    private void loopRows(Closure c) {
        int start = range.firstRow
        int end = range.lastRow
        for (int i = start; i <= end; i++) {
            c.call(CellUtil.getRow(i, sheet), i == start, i == end)
        }
    }

    private void loopColumns(Row row, Closure c) {
        int start = range.firstColumn
        int end = range.lastColumn
        for (int i = start; i <= end; i++) {
            c.call(CellUtil.getCell(row, i), i == start, i == end)
        }
    }

    void setStyles() {
        loopRows { Row row, boolean firstRow, boolean lastRow ->
            loopColumns(row) { Cell cell, boolean firstCol, boolean lastCol ->
                if (firstRow && firstCol) {
                    cell.setCellStyle(leftTop)
                } else if (lastRow && lastCol) {
                    cell.setCellStyle(bottomRight)
                } else {
                    cell.setCellStyle(middle)
                }
            }
        }
    }
}
