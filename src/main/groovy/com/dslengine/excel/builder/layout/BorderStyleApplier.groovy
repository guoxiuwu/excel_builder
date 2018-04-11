package com.dslengine.excel.builder.layout

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide


interface BorderStyleApplier {

    void applyStyle(BorderSide side, BorderStyle style)

    void applyStyle(BorderStyle style)

    void applyColor(BorderSide side, XSSFColor color)

    void applyColor(XSSFColor color)
}
