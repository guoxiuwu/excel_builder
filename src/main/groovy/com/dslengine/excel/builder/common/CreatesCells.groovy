package com.dslengine.excel.builder.common

import com.dslengine.excel.builder.data.Base
import com.dslengine.excel.builder.layout.CellRangeBorderStyleApplier
import org.apache.poi.common.usermodel.Hyperlink
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFHyperlink
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.awt.Color

abstract class CreatesCells implements Base, DoLoop {

    protected final XSSFWorkbook workbook
    protected final XSSFSheet sheet
    protected Map defaultOptions
    protected final Map<Object, Integer> columnIndexes
    protected final CellStyleBuilder styleBuilder
    protected static final Map LINK_OPTIONS = [font: [style: Font.UNDERLINE, color: Color.BLUE]]

    CreatesCells(XSSFSheet sheet, Map defaultOptions, Map<Object, Integer> columnIndexes, CellStyleBuilder styleBuilder) {
        this.workbook = sheet.workbook
        this.sheet = sheet
        this.defaultOptions = defaultOptions
        this.columnIndexes = columnIndexes
        this.styleBuilder = styleBuilder
    }

    /**
     * Sets the default styles to use for the given row or column
     *
     * @param options The style options
     */
    void defaultStyle(Map options) {
        options = new LinkedHashMap(options)
        styleBuilder.convertSimpleOptions(options)
        if (defaultOptions == null) {
            defaultOptions = options
        } else {
            defaultOptions = styleBuilder.merge(defaultOptions, options)
        }
    }

    protected abstract XSSFCell nextCell()

    /**
     * Skips cells
     *
     * @param num The number of cells to skip
     */
    abstract void skipCells(int num)

    protected void setStyle(Object value, XSSFCell cell, Map options) {
        styleBuilder.setStyle(value, cell, options, defaultOptions)
    }

    /**
     * Creates a header cell
     *
     * @param value The cell value
     * @param id The cell identifer
     * @param style The cell style
     * @return The native cell
     */
    XSSFCell column(String value, Object id, final Map style = null) {
        XSSFCell col = cell(value, style)
        columnIndexes[id] = col.columnIndex
        col
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param formulaString The formula
     * @param style The cell style
     * @return The native cell
     */
    XSSFCell formula(String formulaString, final Map style) {
        XSSFCell cell = nextCell()
        if (formulaString.startsWith('=')) {
            formulaString = formulaString[1..-1]
        }
        cell.setCellFormula(formulaString)
        setStyle(null, cell, style)
        cell
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param formulaString The formula
     * @return The native cell
     */
    XSSFCell formula(String formulaString) {
        formula(formulaString, null)
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param callable The return value will be assigned to the cell formula. The closure delegate contains helper methods to get references to other cells.
     * @return The native cell
     */
    XSSFCell formula(@DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellFinder) Closure callable) {
        formula(null, callable)
    }

    /**
     * Assigns a formula to a new cell
     *
     * @param style The cell style
     * @param callable The return value will be assigned to the cell formula. The closure delegate contains helper methods to get references to other cells.
     * @return The native cell
     */
    XSSFCell formula(
            final Map style, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellFinder) Closure callable) {
        XSSFCell cell = nextCell()
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new CellFinder(cell, columnIndexes)
        String formula
        if (callable.maximumNumberOfParameters == 1) {
            formula = (String) callable.call(cell)
        } else {
            formula = (String) callable.call()
        }
        if (formula.startsWith('=')) {
            formula = formula[1..-1]
        }
        cell.setCellFormula(formula)
        setStyle(null, cell, style)
        cell
    }

    /**
     * Creates a new blank cell
     *
     * @return The native cell
     */
    XSSFCell cell() {
        cell(null)
    }

    /**
     * Creates a new cell and assigns a value
     *
     * @param value The value to assign
     * @return The native cell
     */
    XSSFCell cell(Object value) {
        cell(value, null)
    }

    /**
     * Creates a new cell with a value and style
     *
     * @param value The value to assign
     * @param style The cell style options
     * @return The native cell
     */
    XSSFCell cell(Object value, final Map style) {
        XSSFCell cell = nextCell()
        setStyle(value, cell, style)
        if (value == null) {
            return cell
        }
        Closure callable = Excel.getRenderer(value.class)
        if (callable != null) {
            callable.setDelegate(this)
            callable.setResolveStrategy(Closure.DELEGATE_FIRST)
            value = callable.call(value)
        }
        if (value instanceof String) {
            cell.setCellValue((String) value)
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value)
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value)
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue())
        } else if (value instanceof Boolean) {
            cell.setCellValue(value as Boolean)
        } else if (value instanceof Closure) {
            value.setDelegate(this)
            value.setResolveStrategy(Closure.DELEGATE_FIRST)
            def result = value.call()
            if (result instanceof String) {
                cell.setCellValue((String) result)
            } else if (result instanceof Calendar) {
                cell.setCellValue((Calendar) result)
            } else if (result instanceof Date) {
                cell.setCellValue((Date) result)
            } else if (result instanceof Number) {
                cell.setCellValue(((Number) result).doubleValue())
            } else if (result instanceof Boolean) {
                cell.setCellValue(result as Boolean)
            } else {
                cell.setCellValue(result.toString())
            }
        } else {
            cell.setCellValue(value.toString())
        }
        cell
    }

    XSSFCell cell(final Map style, Closure value) {
        XSSFCell cell = nextCell()
        setStyle(value, cell, style)
        if (value == null) {
            return cell
        }
        Closure callable = Excel.getRenderer(value.class)
        if (callable != null) {
            callable.setDelegate(this)
            callable.setResolveStrategy(Closure.DELEGATE_FIRST)
            value = callable.call(value)
        }
        value.setDelegate(this)
        value.setResolveStrategy(Closure.DELEGATE_FIRST)
        def result = value.call()
        if (result instanceof String) {
            cell.setCellValue((String) result)
        } else if (result instanceof Calendar) {
            cell.setCellValue((Calendar) result)
        } else if (result instanceof Date) {
            cell.setCellValue((Date) result)
        } else if (result instanceof Number) {
            cell.setCellValue(((Number) result).doubleValue())
        } else if (result instanceof Boolean) {
            cell.setCellValue(result as Boolean)
        } else {
            cell.setCellValue(result.toString())
        }
        cell
    }

    protected XSSFCell handleLink(XSSFCell cell, String address, HyperlinkType linkType) {
        XSSFHyperlink link = workbook.creationHelper.createHyperlink(linkType)
        link.address = address
        cell.hyperlink = link
        cell
    }

    /**
     * Creates a cell with a hyperlink
     *
     * @param value The cell value
     * @param address The link address
     * @param linkType The type of link. One of {@link Hyperlink#LINK_URL}, {@link Hyperlink#LINK_EMAIL}, {@link Hyperlink#LINK_FILE}
     * @return The native cell
     *
     * @deprecated Use {@link #link(Object, String, HyperlinkType }
     */
    @Deprecated
    XSSFCell link(Object value, String address, int linkType) {
        link(value, address, HyperlinkType.forInt(linkType))
    }

    /**
     * Creates a cell with a hyperlink
     *
     * @param value The cell value
     * @param address The link address
     * @param linkType The type of link. One of {@link HyperlinkType#URL}, {@link HyperlinkType#EMAIL}, {@link HyperlinkType#FILE}
     * @return The native cell
     */
    XSSFCell link(Object value, String address, HyperlinkType linkType) {
        XSSFCell cell = cell(value, LINK_OPTIONS)
        handleLink(cell, address, linkType)
    }

    /**
     * Creates a cell with a hyperlink to another cell in the document
     *
     * @param value The cell value
     * @param callable The closure to build the link
     * @return The native cell
     */
    XSSFCell link(Object value, @DelegatesTo(strategy = Closure.DELEGATE_FIRST, value = CellFinder) Closure callable) {
        XSSFCell cell = cell(value, LINK_OPTIONS)
        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = new CellFinder(cell, columnIndexes)
        handleLink(cell, callable.call().toString(), HyperlinkType.DOCUMENT)
    }

    /**
     * Merges cells
     *
     * @param style Default styles for merged cells
     * @param callable To build cell data
     */
    abstract void merge(final Map style, Closure callable)

    /**
     * Merges cells
     *
     * @param callable To build cell data
     */
    abstract void merge(Closure callable)

    /**
     * Merges cells
     *
     * @param value The cell content
     * @param count How many cells to merge
     * @param style Styling for the merged cell
     */
    void merge(Object value, Integer count, final Map style = null) {
        merge(style) {
            cell(value)
            skipCells(count)
        }
    }

    protected void performMerge(Map style, Closure callable) {
        Map existingDefaultOptions = defaultOptions

        if (style != null && !style.isEmpty()) {
            Map newDefaultOptions = new LinkedHashMap(style)
            styleBuilder.convertSimpleOptions(newDefaultOptions)
            newDefaultOptions = styleBuilder.merge(defaultOptions, newDefaultOptions)
            defaultOptions = newDefaultOptions
        }

        Map borderOptions = defaultOptions?.containsKey('border') ? (Map) defaultOptions.remove('border') : Collections.emptyMap()

        callable.resolveStrategy = Closure.DELEGATE_FIRST
        callable.delegate = this
        int startingIndex = mergeIndex
        callable.call()
        int endingIndex = mergeIndex - 1
        if (endingIndex > startingIndex) {
            CellRangeAddress range = getRange(startingIndex, endingIndex)
            sheet.addMergedRegion(range)
            if (!borderOptions.isEmpty()) {
                styleBuilder.applyBorderToRegion(getBorderStyleApplier(range, sheet), borderOptions)
            }
        }

        defaultOptions = existingDefaultOptions
    }

    protected abstract CellRangeAddress getRange(int start, int end)

    protected abstract int getMergeIndex()

    protected abstract CellRangeBorderStyleApplier getBorderStyleApplier(CellRangeAddress range, XSSFSheet sheet)

}
