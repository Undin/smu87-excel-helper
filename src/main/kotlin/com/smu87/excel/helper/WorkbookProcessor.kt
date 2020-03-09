package com.smu87.excel.helper

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import kotlin.math.roundToInt

// layout:
// number | supplier | name | units | amount | price | cost
class WorkbookProcessor(private val workbook: XSSFWorkbook) {

    private val outputWorkbook: XSSFWorkbook = XSSFWorkbook()

    private val precisionStyle: CellStyle
    private val boldStyle: CellStyle
    private val boldPrecisionStyle: CellStyle

    init {
        val dataFormat = outputWorkbook.createDataFormat()
        val precisionFormatIndex = dataFormat.getFormat("0.00")

        val boldFont = outputWorkbook.createFont()
        boldFont.bold = true

        precisionStyle = outputWorkbook.createCellStyle(null, precisionFormatIndex)
        boldStyle = outputWorkbook.createCellStyle(boldFont, null)
        boldPrecisionStyle = outputWorkbook.createCellStyle(boldFont, precisionFormatIndex)
    }

    fun process(): XSSFWorkbook {
        for (sheet in workbook) {
            val outputSheet = outputWorkbook.createSheet(sheet.sheetName)
            processSheet(sheet, outputSheet)
        }
        return outputWorkbook
    }

    private fun processSheet(sheet: Sheet, outputSheet: Sheet) {
        val titleRow = outputSheet.createRow(0)
        for ((index, columnName) in listOf(
            "№",
            "Поставщик",
            "Наименование",
            "Ед. измерения",
            "Количество",
            "Цена",
            "Сумма"
        ).withIndex()) {
            val cell = titleRow.createCell(index, CellType.STRING)
            cell.setCellValue(columnName)
        }

        sheet
            .filter { !it.getCell(4).value.startsWith("-") }
            .sortedWith(ROW_COMPARATOR)
            .forEachIndexed { index, row ->
                val outputRow = outputSheet.createRow(index + 1)
                if (outputRow.isMainSupplierRow) {
                    outputRow.rowStyle = boldStyle
                }
                processRow(row, outputRow)
            }
    }

    private fun processRow(row: Row, outputRow: Row) {
        var multiplier = 1
        var price: Double? = null
        var amount: Double? = null
        var isCountable = false
        for (i in 0 until 7) {
            // number, amount, price and cost are numbers. otherwise it's string
            val cellType = if (i in listOf(0, 4, 5, 6)) CellType.NUMERIC else CellType.STRING
            val outputCell = outputRow.createCell(i, cellType)
            var setPrecisionStyle = false

            when (i) {
                // number
                0 -> {
                    outputCell.setCellValue(outputRow.rowNum.toDouble())
                }
                // unit
                3 -> {
                    val unit = row[i].value
                    val result = UNIT_PATTERN.matchEntire(unit)?.groups
                    val unitValue = if (result != null) {
                        multiplier = result["count"]!!.value.toInt()
                        result["unit"]!!.value
                    } else {
                        unit
                    }
                    if ("шт" in unitValue) {
                        isCountable = true
                    }
                    outputCell.setCellValue(unitValue)
                }
                // amount
                4 -> {
                    val value = row[i].value
                    amount = value.substringBefore(" x").replace(",", ".").toDouble() * multiplier
                    if (isCountable) {
                        amount = amount.roundToInt().toDouble()
                    }
                    outputCell.setCellValue(amount)
                }
                // price
                5 -> {
                    price = row[i].value.toDouble() / multiplier
                    outputCell.setCellValue(price)
                    setPrecisionStyle = true
                }
                // cost
                6 -> {
                    outputCell.setCellValue(price!! * amount!!)
                    setPrecisionStyle = true
                }
                else -> {
                    outputCell.setCellValue(row[i].value)
                }
            }

            outputCell.cellStyle = when {
                row.isMainSupplierRow -> if (setPrecisionStyle) boldPrecisionStyle else boldStyle
                setPrecisionStyle -> precisionStyle
                else -> null
            }
        }
    }

    companion object {
        private val SECONDARY_SUPPLIER: Regex = """\d*-\d*""".toRegex()
        private val UNIT_PATTERN: Regex = """(?<count>\d+) (?<unit>.+)""".toRegex()
        private val ROW_COMPARATOR: Comparator<Row> = Comparator.comparing<Row, Boolean> {
            !it.isMainSupplierRow
        }.thenBy {
            it.getCell(2).value
        }

        private val Row.isMainSupplierRow: Boolean get() = !getCell(1).value.matches(SECONDARY_SUPPLIER)

        private operator fun Row.get(index: Int): Cell? = getCell(index)

        private val Cell?.value: String get() = this?.toString().orEmpty().trim()

        private fun XSSFWorkbook.createCellStyle(font: Font?, dataFormatIndex: Short?): CellStyle {
            val style = createCellStyle()
            if (font != null) {
                style.setFont(font)
            }
            if (dataFormatIndex != null) {
                style.dataFormat = dataFormatIndex
            }
            return style
        }
    }
}