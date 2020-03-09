package com.smu87.excel.helper

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import kotlin.math.roundToInt

// layout:
// number | supplier | name | units | amount | price | cost
class WorkbookProcessor(private val workbook: XSSFWorkbook) {

    private val outputWorkbook: XSSFWorkbook = XSSFWorkbook()

    private val precisionStyle: CellStyle

    init {
        precisionStyle = outputWorkbook.createCellStyle()
        val dataFormat = outputWorkbook.createDataFormat()
        val precisionFormat = dataFormat.getFormat("0.00")
        precisionStyle.dataFormat = precisionFormat
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
                processRow(row, outputRow)
            }
    }

    private fun processRow(row: Row, outputRow: Row) {
        var multiplier = 1
        var price: Double? = null
        var amount: Double? = null
        var isCountable = false
        for (i in 0 until 7) {
            when (i) {
                // number
                0 -> {
                    val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                    outputCell.setCellValue(outputRow.rowNum.toDouble())
                }
                // unit
                3 -> {
                    val outputCell = outputRow.createCell(i, CellType.STRING)
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
                    val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                    outputCell.setCellValue(amount)
                }
                // price
                5 -> {
                    price = row[i].value.toDouble() / multiplier
                    val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                    outputCell.setCellValue(price)
                    outputCell.cellStyle = precisionStyle
                }
                // cost
                6 -> {
                    val outputCell = outputRow.createCell(i, CellType.NUMERIC)
                    outputCell.setCellValue(price!! * amount!!)
                    outputCell.cellStyle = precisionStyle
                }
                else -> {
                    val outputCell = outputRow.createCell(i, CellType.STRING)
                    outputCell.setCellValue(row[i].value)
                }
            }
        }
    }

    companion object {
        private val SECONDARY_SUPPLIER: Regex = """\d*-\d*""".toRegex()
        private val UNIT_PATTERN: Regex = """(?<count>\d+) (?<unit>.+)""".toRegex()
        private val ROW_COMPARATOR: Comparator<Row> = Comparator.comparing<Row, Boolean> {
            it.getCell(1).value.matches(SECONDARY_SUPPLIER)
        }.thenBy {
            it.getCell(2).value
        }

        private operator fun Row.get(index: Int): Cell? = getCell(index)

        private val Cell?.value: String get() = this?.toString().orEmpty().trim()
    }
}