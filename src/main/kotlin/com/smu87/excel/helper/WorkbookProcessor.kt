package com.smu87.excel.helper

import org.apache.logging.log4j.LogManager
import org.apache.logging.log4j.Logger
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import kotlin.math.roundToInt

// layout:
// number | supplier | name | units | amount | price | cost
class WorkbookProcessor(private val workbook: XSSFWorkbook) {

    private val outputWorkbook: XSSFWorkbook = XSSFWorkbook()

    private val commonStyle: CellStyle
    private val precisionStyle: CellStyle
    private val boldStyle: CellStyle
    private val boldPrecisionStyle: CellStyle

    init {
        val dataFormat = outputWorkbook.createDataFormat()
        val precisionFormatIndex = dataFormat.getFormat("0.00")

        val commonFont = outputWorkbook.createFont()
        commonFont.fontName = FONT_NAME
        commonFont.setFontHeight(FONT_SIZE)

        val boldFont = outputWorkbook.createFont()
        boldFont.bold = true
        boldFont.fontName = FONT_NAME
        boldFont.setFontHeight(FONT_SIZE)

        commonStyle = outputWorkbook.createCellStyle(commonFont, null)
        precisionStyle = outputWorkbook.createCellStyle(commonFont, precisionFormatIndex)
        boldStyle = outputWorkbook.createCellStyle(boldFont, null)
        boldPrecisionStyle = outputWorkbook.createCellStyle(boldFont, precisionFormatIndex)
    }

    fun process(): XSSFWorkbook {
        LOG.info("Start processing. Sheets: ${workbook.numberOfSheets}")
        for (sheet in workbook) {
            val outputSheet = outputWorkbook.createSheet(sheet.sheetName)
            try {
                processSheet(sheet, outputSheet)
            } catch (e: Exception) {
                LOG.error(e.message, e)
            }
        }
        LOG.info("End processing")
        return outputWorkbook
    }

    private fun processSheet(sheet: Sheet, outputSheet: Sheet) {
        LOG.info("Start processing of `${sheet.sheetName}` sheet")
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
            cell.cellStyle = commonStyle
        }

        val materials = collectMaterialInfo(sheet)

        for ((index, info) in materials.withIndex()) {
            val outputRow = outputSheet.createRow(index + 1)
            if (info.material.isFromMainSupplier) {
                outputRow.rowStyle = boldStyle
            }
            val isLastMainSupplier = info.material.isFromMainSupplier &&
                    (materials.getOrNull(index + 1)?.material?.isFromMainSupplier?.not() ?: true)

            writeToRow(info, outputRow, writeSum = isLastMainSupplier)
        }

        LOG.info("End processing of `${sheet.sheetName}` sheet")
    }

    private fun collectMaterialInfo(sheet: Sheet): List<MaterialInfo> {
        return sheet
            .filter { !it.getCell(4).value.startsWith("-") } // negative amount should be ignored
            .map { collectMaterialInfo(it) }
            .filter { it.cost != 0.0 }
            .groupBy { it.material }
            .map { (_, materials) -> materials.reduce { acc, info -> acc + info } }
            .sortedBy { it.material }
    }

    private fun collectMaterialInfo(row: Row): MaterialInfo {
        val builder = MaterialInfoBuilder()
        var multiplier = 1
        var isCountable = false

        for (i in 1 until 6) {
            when (i) {
                1 -> builder.supplier = row[i].value
                2 -> builder.name = row[i].value
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
                    builder.units = unitValue
                }
                // amount
                4 -> {
                    val value = row[i].value
                    var amount = value.substringBefore(" x").replace(",", ".").toDouble() * multiplier
                    if (isCountable) {
                        amount = amount.roundToInt().toDouble()
                    }
                    builder.amount = amount
                }
                // price
                5 -> builder.price = row[i].value.toDouble() / multiplier
            }
        }

        return builder.build()
    }

    private fun writeToRow(info: MaterialInfo, outputRow: Row, writeSum: Boolean) {
        val rowNum = outputRow.rowNum + 1
        for (i in 0 until 7) {
            val (value, setPrecisionStyle) = when (i) {
                // number
                0 -> Number(outputRow.rowNum.toDouble()) to false
                // supplier
                1 -> Str(info.material.supplier) to false
                // material name
                2 -> Str(info.material.name) to false
                // unit
                3 -> Str(info.material.units) to false
                // amount
                4 -> Number(info.amount) to false
                // price
                5 -> Number(info.material.price) to true
                // cost
                6 -> Formula("ROUND(E$rowNum*F$rowNum,2)") to true
                else -> error("Unexpected cell index: $i")
            }

            val outputCell = outputRow.createCell(i, value.cellType)

            when (value) {
                is Number -> outputCell.setCellValue(value.value)
                is Str -> outputCell.setCellValue(value.value)
                is Formula -> outputCell.cellFormula = value.value
            }

            outputCell.cellStyle = when {
                info.material.isFromMainSupplier -> if (setPrecisionStyle) boldPrecisionStyle else boldStyle
                setPrecisionStyle -> precisionStyle
                else -> commonStyle
            }

            if (writeSum) {
                val sumCell = outputRow.createCell(7)
                sumCell.cellFormula = "SUM(G2:G$rowNum)"
                sumCell.cellStyle = boldPrecisionStyle
            }
        }
    }

    companion object {

        private val LOG: Logger = LogManager.getLogger(WorkbookProcessor::class.java)

        private const val FONT_NAME: String = "Times New Roman"
        private const val FONT_SIZE: Double = 10.0

        private val UNIT_PATTERN: Regex = """(?<count>\d+) (?<unit>.+)""".toRegex()

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

    private sealed class Value {
        val cellType: CellType get() {
            return when (this) {
                is Number -> CellType.NUMERIC
                is Str -> CellType.STRING
                is Formula -> CellType.FORMULA
            }
        }
    }
    private class Number(val value: Double) : Value()
    private class Str(val value: String) : Value()
    private class Formula(val value: String) : Value()

}
