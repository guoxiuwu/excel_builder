package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.data.Base
import groovy.transform.CompileStatic
import org.apache.poi.ss.util.WorkbookUtil
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class WorkBook implements Base, DoLoop {

    private final XSSFWorkbook wb
    private final CellStyleBuilder styleBuilder

    private static final String WIDTH = 'width'
    private static final String HEIGHT = 'height'

    def call(Closure cl) {
        cl.setDelegate(this)
        cl.setResolveStrategy(Closure.DELEGATE_FIRST)
        cl.call()
    }

    WorkBook(XSSFWorkbook wb) {
        this.wb = wb
        this.styleBuilder = new CellStyleBuilder(wb)
    }

    /**
     * Creates a sheet
     *
     * @param name The sheet name
     * @param options Default sheet options
     * @param callable To build data
     * @return The native sheet object
     */
    XSSFSheet sheet(String name, Map options,
                    @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(WorkbookUtil.createSafeSheetName(name)), options, callable)
    }

    /**
     * Creates a sheet
     *
     * @param name The sheet name
     * @param callable To build data
     * @return The native sheet object
     */
    XSSFSheet sheet(String name, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet(name, [:], callable)
    }

    /**
     * Creates a sheet
     *
     * @param callable To build data
     * @return The native sheet object
     */
    XSSFSheet sheet(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        sheet([:], callable)
    }

    /**
     * Creates a sheet
     *
     * @param options Default sheet options
     * @param callable To build data
     * @return The native sheet object
     */
    XSSFSheet sheet(Map options, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = Sheet) Closure callable) {
        handleSheet(wb.createSheet(), options, callable)
    }

    private XSSFSheet handleSheet(XSSFSheet sheet, Map options, Closure callable) {
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        if (options.containsKey(WIDTH)) {
            def width = options[WIDTH]
            if (width instanceof Integer) {
                sheet.setDefaultColumnWidth(width)
            } else {
                throw new IllegalArgumentException('Sheet default column width must be an integer')
            }
        }

        if (options.containsKey(HEIGHT)) {
            def height = options[HEIGHT]
            if (height instanceof Short) {
                sheet.setDefaultRowHeight(height)
            } else if (height instanceof Float) {
                sheet.setDefaultRowHeightInPoints(height)
            } else {
                throw new IllegalArgumentException('Sheet default row height must be a short or float')
            }
        }

        callable.delegate = new Sheet(sheet, styleBuilder)
        if (callable.maximumNumberOfParameters == 1) {
            callable.call(sheet)
        } else {
            callable.call()
        }
        sheet
    }
}
