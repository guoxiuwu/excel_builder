package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.data.DataContext
import com.dslengine.excel.builder.data.ExcelBuilderContext
import groovy.transform.CompileStatic
import org.apache.poi.xssf.usermodel.XSSFWorkbook

@CompileStatic
class ExcelBuilder {

    private static final String dslStartBlockName = "excel"

    static void buildExcel(OutputStream outputStream, String dslFileAbsPath, DataContext dataContext) {
        XSSFWorkbook wb = buildExcel(new File(dslFileAbsPath), dataContext)
        wb.write(outputStream)
    }

    static void buildExcel(OutputStream outputStream, File dslFile, DataContext dataContext) {
        XSSFWorkbook wb = buildExcel(dslFile, dataContext)
        wb.write(outputStream)
    }

    static XSSFWorkbook buildExcel(String dslFileAbsPath, DataContext dataContext) {
        buildExcel(new File(dslFileAbsPath), dataContext)
    }

    static void buildExcel(String dslFileAbsPath, String outputExcelPath, DataContext dataContext) {
        buildExcel(new File(dslFileAbsPath), new File(outputExcelPath), dataContext)
    }

    static void buildExcel(File dslFile, File outputExcelFile, DataContext dataContext) {
        OutputStream oi
        try {
            oi = new FileOutputStream(outputExcelFile)
            XSSFWorkbook wb = buildExcel(dslFile, dataContext)
            wb.write(oi)
        } finally {
            oi?.close()
        }
    }

    static XSSFWorkbook buildExcel(File dslFile, DataContext dataContext) {
        if (!dslFile.exists()) {
            throw new FileNotFoundException("dslFile:${dslFile.absolutePath} not found")
        }
        XSSFWorkbook wb
        try {
            wb = new XSSFWorkbook()
            GroovyShell groovyShell = new GroovyShell()
            ExcelBuilderContext.initDataContext(dataContext)
            groovyShell.setVariable(dslStartBlockName, new WorkBook(wb))
            groovyShell.run(dslFile, [] as String[])
        } finally {
            ExcelBuilderContext.releaseDataContext()
        }
        wb
    }

}
