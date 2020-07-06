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
            writeToRow(info, outputRow)
        }

        LOG.info("End processing of `${sheet.sheetName}` sheet")
    }

    private fun collectMaterialInfo(sheet: Sheet): List<MaterialInfo> {
        val (fromMainSuppliers, fromSecondarySuppliers) = sheet
            .filter { !it.getCell(4).value.startsWith("-") } // negative amount should be ignored
            .map { collectMaterialInfo(it) }
            .filter { it.cost != 0.0 }
            .partition { it.material.isFromMainSupplier }

        val collapsedFromSecondarySuppliers = fromSecondarySuppliers
            .groupBy { it.material }
            .map { (_, materials) -> materials.reduce { acc, info -> acc + info } }

        return (fromMainSuppliers + collapsedFromSecondarySuppliers).sortedBy { it.material }
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

    private fun writeToRow(info: MaterialInfo, outputRow: Row) {
        for (i in 0 until 7) {
            val (value, setPrecisionStyle) = when (i) {
                // number
                0 -> outputRow.rowNum.toDouble() to false
                // supplier
                1 -> info.material.supplier to false
                // material name
                2 -> info.material.name to false
                // unit
                3 -> info.material.units to false
                // amount
                4 -> info.amount to false
                // price
                5 -> info.material.price to true
                // cost
                6 -> info.cost to true
                else -> error("Unexpected cell index: $i")
            }

            val cellType = if (value is Double) CellType.NUMERIC else CellType.STRING
            val outputCell = outputRow.createCell(i, cellType)
            when (value) {
                is Double -> outputCell.setCellValue(value)
                is String -> outputCell.setCellValue(value)
                else -> error("Unexpected value type: ${value.javaClass}")
            }

            outputCell.cellStyle = when {
                info.material.isFromMainSupplier -> if (setPrecisionStyle) boldPrecisionStyle else boldStyle
                setPrecisionStyle -> precisionStyle
                else -> commonStyle
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
}
